"""متانة: تغيّر شكل الملفات (ترتيب أعمدة، عناوين، صفوف زائدة، أرقام نصية...) لا يكسر البرنامج
ولا يُنقص جودة المطابقة بشكل ملحوظ."""
import csv
import random
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl
import pytest

from app.analytics import analyze
from app.engine import reconcile, refresh_result
from app.exporter import export_report
from app.loader import auto_map, build_records, load_table, mapping_problems
from app.sample import make


def _sheet_rows(path):
    ws = openpyxl.load_workbook(path).active
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _find_header(rows):
    best, bi = -1, 0
    for i, r in enumerate(rows[:12]):
        n = sum(1 for c in r if isinstance(c, str) and c.strip())
        if n > best:
            best, bi = n, i
    return bi


def _write(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    wb.save(path)
    return str(path)


def _run(pa, pb):
    ta, tb = load_table(pa), load_table(pb)
    ma, mb = auto_map(ta), auto_map(tb)
    A, _ = build_records(ta, ma, "A")
    B, _ = build_records(tb, mb, "B")
    res = reconcile(A, B)
    refresh_result(res)
    analyze(res)
    return res, ta, tb


@pytest.fixture(scope="module")
def base():
    a, b = make(tempfile.mkdtemp())
    res, *_ = _run(a, b)
    return a, b, res.counts["متطابقة"]


def _variant(path, out, fn):
    rows = _sheet_rows(path)
    h = _find_header(rows)
    return _write(out, fn(rows, h))


def shuffle_cols(rows, h):
    n = max(len(r) for r in rows)
    perm = list(range(n))
    random.Random(3).shuffle(perm)
    return [[(r + [None] * n)[j] for j in perm] for r in rows]


def extra_rows(rows, h):
    body = rows[h + 1:]
    mid = len(body) // 2
    junk = [None] * len(rows[h])
    return ([["تقرير مصدّر", "نسخة 2"], [], ["", ""]] + rows[:h + 1] + body[:mid] + [junk, junk] + body[mid:]
            + [junk, ["الإجمالي"] + [None] * (len(rows[h]) - 2) + [123456]])


def text_numbers(rows, h):
    out = []
    for i, r in enumerate(rows):
        if i <= h:
            out.append(r)
            continue
        nr = []
        for c in r:
            if isinstance(c, (int, float)) and not isinstance(c, bool):
                nr.append(f"{c:,}".translate(str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")))
            else:
                nr.append(c)
        out.append(nr)
    return out


def padded_headers(rows, h):
    rows = [list(r) for r in rows]
    rows[h] = [(f"  {c}  " if isinstance(c, str) else c) for c in rows[h]]
    return [r + [None, None] for r in rows]


def drop_col(name_part):
    def fn(rows, h):
        j = next((k for k, c in enumerate(rows[h]) if isinstance(c, str) and name_part in c.lower()), None)
        if j is None:
            return rows
        return [[c for k, c in enumerate(r) if k != j] for r in rows]
    return fn


@pytest.mark.parametrize("name,fn", [
    ("shuffle", shuffle_cols), ("extra_rows", extra_rows), ("text_numbers", text_numbers),
    ("padded", padded_headers),
])
def test_layout_changes_keep_quality(base, tmp_path, name, fn):
    a, b, n0 = base
    for src, tag in ((a, "a"), (b, "b")):
        pass
    pa = _variant(a, tmp_path / "a.xlsx", fn)
    pb = _variant(b, tmp_path / "b.xlsx", fn)
    res, ta, tb = _run(pa, pb)
    assert res.counts["متطابقة"] >= 0.9 * n0
    assert all(v == 0 for v in res.integrity.values())
    export_report(str(tmp_path / "r.xlsx"), res, ta, tb)


@pytest.mark.parametrize("col", ["date", "commission", "beneficiary", "currency"])
def test_missing_columns_do_not_crash(base, tmp_path, col):
    a, b, n0 = base
    pb = _variant(b, tmp_path / "b.xlsx", drop_col(col))
    res, ta, tb = _run(a, pb)
    assert res.counts["متطابقة"] >= 0.5 * n0
    assert all(v == 0 for v in res.integrity.values())
    export_report(str(tmp_path / "r.xlsx"), res, ta, tb)


def test_csv_variant(base, tmp_path):
    a, b, n0 = base
    rows = _sheet_rows(b)
    p = tmp_path / "b.csv"
    with open(p, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(rows)
    res, *_ = _run(a, str(p))
    assert res.counts["متطابقة"] >= 0.9 * n0


def test_garbage_does_not_crash(tmp_path):
    rows = [["x", "y"], [None, None], ["abc", "def"], [1, 2], ["", ""], ["٣", "٤"]]
    p = _write(tmp_path / "g.xlsx", rows)
    t = load_table(p)
    m = auto_map(t)
    recs, skipped = build_records(t, m, "A")        # قد لا تُنتج سجلات لكنها لا تنهار
    assert isinstance(recs, list)
    assert isinstance(mapping_problems(m), list)


def test_empty_one_side(base):
    a, b, _ = base
    ta = load_table(a)
    A, _ = build_records(ta, auto_map(ta), "A")
    res = reconcile(A, [])
    assert res.counts["مفقودة في الصرافة"] + res.counts["مكررة"] == len(A)
    analyze(res)


def test_engine_fuzz_never_crashes_and_integrity_closes():
    """سجلات عشوائية قاسية: مبالغ صفرية/ضخمة/سالبة، نصوص غريبة، تواريخ ناقصة، عملات مختلطة."""
    from datetime import date, timedelta
    from decimal import Decimal as D
    from app.loader import Record
    rnd = random.Random(7)
    words = ["حوالة", "باسم", "مناولة", "س.ص", "رقم", "خدمة", "عمولة", "١٢٣٤", "ABC-12", "", "  ", "ـــ", "😀", "x" * 80]
    cur = ["", "SAR", "USD", "YER"]
    for it in range(60):
        def mk(side, n):
            out = []
            for i in range(n):
                amt = D(rnd.choice([0, 1, -1, 5, 100, 100, 12345, -12345, 10**9])) / rnd.choice([1, 1, 100])
                d = date(2026, 9, 1) + timedelta(days=rnd.randint(0, 40)) if rnd.random() > .15 else None
                ref = rnd.choice(["", "", "A1", "70001", "70002"]) 
                desc = " ".join(rnd.choice(words) + (str(rnd.randint(1, 99999)) if rnd.random() < .3 else "") for _ in range(rnd.randint(0, 6)))
                out.append(Record(side, i + 2, ref, ref.upper(), "", d, amt, rnd.choice(cur), D(rnd.choice([0, 0, 5])),
                                  rnd.choice(["", "علي", "سالم"]), desc, raw={"البيان": desc}, roles={"desc": "البيان"}))
            return out
        A, B = mk("A", rnd.randint(0, 25)), mk("B", rnd.randint(0, 25))
        res = reconcile(A, B)
        refresh_result(res)
        analyze(res)
        assert all(v == 0 for v in res.integrity.values()), (it, res.integrity)
