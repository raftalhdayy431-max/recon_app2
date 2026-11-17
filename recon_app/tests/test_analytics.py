"""اختبارات التحليلات: الانكشاف، التصنيفات، الأطراف، الملاحظات، ورقة Excel."""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl

from app.analytics import analyze
from app.engine import reconcile
from app.exporter import export_report
from app.loader import auto_map, build_records, load_table
from app.sample import make


def _res():
    a, b = make(tempfile.mkdtemp())
    ta, tb = load_table(a), load_table(b)
    A, _ = build_records(ta, auto_map(ta), "A")
    B, _ = build_records(tb, auto_map(tb), "B")
    return reconcile(A, B), ta, tb


def test_analysis_basic():
    res, _, _ = _res()
    an = analyze(res)
    assert an.issues_total == sum(v for k, v in res.counts.items() if k != "متطابقة")
    assert an.categories and an.exposure and an.parties and an.findings
    assert sum(t[1] for t in an.timeline) <= an.rows_total
    assert all(c.amount >= 0 for c in an.categories)


def test_analysis_all_matched_gives_good_finding():
    res, _, _ = _res()
    res.rows = [r for r in res.rows if r.status == "متطابقة"]
    an = analyze(res)
    assert an.issues_total == 0 and an.findings[0].level == "good"


def test_report_has_analysis_sheet(tmp_path):
    res, ta, tb = _res()
    p = str(tmp_path / "r.xlsx")
    export_report(p, res, ta, tb)
    wb = openpyxl.load_workbook(p)
    assert "التحليل" in wb.sheetnames
    ws = wb["التحليل"]
    texts = [str(c.value) for row in ws.iter_rows() for c in row if c.value]
    assert any("ملاحظات ذكية" in t for t in texts)


def test_opposite_sign_precaution():
    from datetime import date
    from decimal import Decimal as D
    from app.loader import Record

    def R(side, row, ref, amt, party="زيد", d=date(2026, 9, 1)):
        return Record(side, row, ref, ref.upper(), "", d, D(amt), "SAR", D(0), party, "")
    A = [R("A", i, f"10000{i}", 1000 + i) for i in range(1, 8)] + [R("A", 8, "100008", "500"), R("A", 9, "", "700", "عمرو")]
    B = [R("B", i, f"10000{i}", 1000 + i) for i in range(1, 8)] + [R("B", 8, "100008", "-500"), R("B", 9, "", "-700", "عمرو")]
    res = reconcile(A, B)
    assert all(r.status != "متطابقة" for r in res.rows if r.ref == "100008")   # لا تُعتمد تلقائياً
    an = analyze(res)
    assert any("إشارة معكوسة" in f.title for f in an.findings)


def test_approve_moves_to_matched_and_back(tmp_path):
    from app.engine import refresh_result, set_decision
    res, ta, tb = _res()
    row = next(r for r in res.rows if r.status in ("مختلفة", "محتملة المطابقة") and r.a and r.b)
    before = dict(res.counts)
    st = row.status
    set_decision(row, "معتمد")
    refresh_result(res)
    assert row.status == "متطابقة" and res.counts["متطابقة"] == before["متطابقة"] + 1
    assert res.counts[st] == before[st] - 1
    assert all(v == 0 for v in res.integrity.values())          # الجسر الحسابي يبقى مغلقاً
    p = str(tmp_path / "r.xlsx")
    export_report(p, res, ta, tb)
    ws = openpyxl.load_workbook(p)["المتطابقات"]
    texts = [str(c.value) for rw in ws.iter_rows() for c in rw if c.value]
    assert "اعتماد المراجع" in texts
    set_decision(row, "")
    refresh_result(res)
    assert row.status == st and res.counts == before


def test_manual_link_suggest_and_approve(tmp_path):
    from datetime import date
    from decimal import Decimal as D
    from app.engine import suggest_links, manual_link
    from app.loader import Record

    def R(side, row, ref, amt, desc, d=date(2026, 9, 1)):
        return Record(side, row, ref, ref.upper(), "", d, D(amt), "SAR", D(0), "", desc)
    A = [R("A", i, f"10000{i}", 1000 + i, f"حوالة باسم عميل{i} رقم 10000{i}") for i in range(1, 8)]
    B = [R("B", i, f"10000{i}", -(1000 + i), f"مناولة عميل{i} الرقم العام 10000{i}") for i in range(1, 8)]
    A.append(R("A", 8, "", "777", "دفعة سالم الخولاني"))
    B.append(R("B", 8, "", "-777", "تحويل سالم الخولاني نقدا", date(2026, 9, 9)))   # فارق 6 أيام: لا تُطابَق آلياً
    res = reconcile(A, B)
    ra = next(r for r in res.rows if r.status == "مفقودة في الصرافة")
    rb = next(r for r in res.rows if r.status == "مفقودة في ملف المحاسب")
    sug = suggest_links(res, ra)
    assert sug and sug[0][0] is rb and sug[0][1] >= 50
    n = res.counts["متطابقة"]
    new = manual_link(res, ra, rb, approve=True)
    assert new.status == "متطابقة" and new.decision == "معتمد"
    assert res.counts["متطابقة"] == n + 1 and res.counts["مفقودة في الصرافة"] == 0
    assert all(v == 0 for v in res.integrity.values())


def test_selective_export(tmp_path):
    import pytest
    from app.exporter import SHEET_NAME
    res, ta, tb = _res()
    p = str(tmp_path / "m.xlsx")
    export_report(p, res, ta, tb, sheets=["matched"])
    assert openpyxl.load_workbook(p).sheetnames == [SHEET_NAME["matched"]]
    p2 = str(tmp_path / "s.xlsx")
    export_report(p2, res, ta, tb, sheets=["summary", "review"])
    wb = openpyxl.load_workbook(p2)
    assert wb.sheetnames == [SHEET_NAME["summary"], SHEET_NAME["review"]]
    with pytest.raises(ValueError):
        export_report(str(tmp_path / "x.xlsx"), res, ta, tb, sheets=[])


def test_amount_only_pairs():
    from datetime import date
    from decimal import Decimal as D
    from app.engine import amount_only_pairs, manual_link
    from app.loader import Record

    def R(side, row, amt, desc, d=date(2026, 9, 1)):
        return Record(side, row, "", "", "", d, D(amt), "SAR", D(0), "", desc)
    A = [R("A", 1, "500", "x"), R("A", 2, "321", "y", date(2026, 9, 3))]
    B = [R("B", 1, "-500", "z", date(2026, 9, 20)), R("B", 2, "-321", "w", date(2026, 9, 25))]
    res = reconcile(A, B)
    pairs = amount_only_pairs(res)
    assert len(pairs) == 2 and all(p["unique"] for p in pairs)
    for p in pairs:
        manual_link(res, p["a"], p["b"], approve=True)
    assert res.counts["متطابقة"] == 2 and amount_only_pairs(res) == []
    assert all(v == 0 for v in res.integrity.values())
