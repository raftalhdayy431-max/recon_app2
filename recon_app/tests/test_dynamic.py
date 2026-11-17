"""اختبارات الديناميكية: أعمدة متغيرة، مطابقة بالبيان، قواعد مخصصة، تقرير ديناميكي، PDF."""
import sys
from pathlib import Path
from decimal import Decimal

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl

from app.engine import Settings, reconcile, STATUS_MATCHED, STATUS_MISSING_A, STATUS_MISSING_B
from app.exporter import export_report
from app.keys import KeyRule
from app.loader import auto_map, build_records, load_table


def _xlsx(path, header, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return str(path)


def _recs(tmp_path, ha, ra, hb, rb):
    pa = _xlsx(tmp_path / "a.xlsx", ha, ra)
    pb = _xlsx(tmp_path / "b.xlsx", hb, rb)
    ta, tb = load_table(pa), load_table(pb)
    A, _ = build_records(ta, auto_map(ta), "A")
    B, _ = build_records(tb, auto_map(tb), "B")
    return A, B, ta, tb


def test_match_by_description_without_reference_columns(tmp_path):
    A, B, *_ = _recs(
        tmp_path,
        ["التاريخ", "البيان", "المبلغ"],
        [["2026-01-05", "حوالة رقم 7001234 إلى أحمد", 1000], ["2026-01-06", "حوالة رقم 7005678 إلى سعيد", 250.5],
         ["2026-01-07", "حوالة رقم 7009999 إلى علي", 80]],
        ["Date", "Description", "Amount"],
        [["2026-01-05", "TRF 7001234 AHMED", 1000], ["2026-01-06", "TRF 7005678 SAEED", 250.5],
         ["2026-01-09", "TRF 7004444 OMAR", 40]])
    assert not any(r.ref for r in A + B)             # لا عمود مرجع في الملفين
    res = reconcile(A, B)
    assert [k.name for k in res.key_rules][0] == "البيان"
    assert res.counts[STATUS_MATCHED] == 2
    assert res.counts[STATUS_MISSING_B] == 1 and res.counts[STATUS_MISSING_A] == 1
    assert all(v == 0 for v in res.integrity.values())
    m = [r for r in res.rows if r.status == STATUS_MATCHED]
    assert all("مشتركة بين عمود" in r.reason_text for r in m)
    assert all(r.method == "البيان" for r in m)


def test_fallback_to_reference_when_no_description(tmp_path):
    A, B, *_ = _recs(
        tmp_path,
        ["رقم الحوالة", "المبلغ"], [["A1001", 10], ["A1002", 20]],
        ["Reference", "Amount"], [["a1001", 10], ["A1002", 21]])
    res = reconcile(A, B)
    assert [k.name for k in res.key_rules] == ["المرجع"]
    assert res.counts[STATUS_MATCHED] == 1


def test_other_columns_used_when_nothing_standard(tmp_path):
    A, B, *_ = _recs(
        tmp_path,
        ["كود الطلب", "المبلغ"], [["ORD-778899", 500], ["ORD-112233", 90]],
        ["Order Code", "Amount"], [["ORD-778899", 500], ["ORD-000111", 90]])
    res = reconcile(A, B)
    assert [k.name for k in res.key_rules] == ["بقية الأعمدة"]
    assert res.counts[STATUS_MATCHED] == 1


def test_custom_rules_priority_and_missing_column_warning(tmp_path):
    A, B, *_ = _recs(
        tmp_path,
        ["رقم الحوالة", "البيان", "المبلغ"], [["111111", "ملاحظة 999888", 5]],
        ["Reference", "Memo", "Amount"], [["999888", "x", 5]])
    S = Settings(key_rules=[KeyRule("مخصصة", ["رقم الحوالة", "عمود_غير_موجود"], ["Reference"], "exact").to_dict(),
                            KeyRule("ثانية", ["البيان"], ["Reference"], "text").to_dict()])
    res = reconcile(A, B, S)
    assert any("غير موجود" in w for w in res.warnings)
    assert [k.name for k in res.key_rules] == ["مخصصة", "ثانية"]
    # القاعدة الأولى لا تطابق (111111 ≠ 999888)، فتطابق الثانية عبر البيان
    assert res.rows[0].method == "ثانية" and res.rows[0].status == STATUS_MATCHED


def test_report_is_dynamic(tmp_path):
    A, B, ta, tb = _recs(
        tmp_path,
        ["التاريخ", "البيان", "المبلغ"], [["2026-01-05", "حوالة 7001234", 100]],
        ["Date", "Details", "Amount"], [["2026-01-05", "TRF 7001234", 100]])
    for allc in (True, False):
        res = reconcile(A, B, Settings(report_all_columns=allc))
        out = tmp_path / f"r{allc}.xlsx"
        export_report(str(out), res, ta, tb)
        ws = openpyxl.load_workbook(out)["جدول المطابقة"]
        hdr = [c.value for c in ws[4]]
        assert "المرجع" not in hdr and "المرجع الثانوي" not in hdr and "الرسوم" not in hdr
        assert "البيان" in hdr or "Details" in hdr


def test_pdf_table_reading(tmp_path):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    p = tmp_path / "t.pdf"
    data = [["Reference", "Date", "Amount", "Description"]] + \
           [[f"70012{i}", f"2026-01-0{i}", f"{100 * i}.00", f"payment {i}"] for i in range(1, 6)]
    t = Table(data)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    SimpleDocTemplate(str(p), pagesize=A4).build([t])
    tb = load_table(str(p))
    assert tb.columns[:3] == ["Reference", "Date", "Amount"]
    assert len(tb.rows) == 5
    recs, _ = build_records(tb, auto_map(tb), "A")
    assert len(recs) == 5 and recs[0].amount == Decimal("100.00")


def test_real_world_layout_voucher_and_names(tmp_path):
    """نموذج واقعي: عمودا مدين/دائن بعنوان «المبلغ مدين»، رقم مستند قصير يُذكر في بيان الصرافة،
    وحوالات بلا مرجع مشترك تُطابَق بالمبلغ والاسم الموجود في البيان."""
    import datetime as dt
    d = dt.datetime(2026, 9, 1)
    ha = ["المبلغ دائن", "المبلغ مدين", "الـــــتــفاصــــيل", "رقم المستند", "نوع السند", "التاريخ"]
    ra = [
        [85000, 0, "امين صالح حسين المرهبي - حارس عصر", 3691, "سند صرف نقدي", d],
        [0, 209520, "حوالة من بشير عتيق الحزوره الرقم العام 5570371718", 167, "سند قبض نقدي", d],
        [25000, 0, "حواله باسم جلال قاسم المريسي رقم 5027973653", 3697, "سند صرف نقدي", d],
        [25000, 0, "حوالة باسم صدام حسين الاديمي برقم 1122499475", 3719, "سند صرف نقدي", d],
        [109000, 0, "استثمار تصفية عهدة محمد شويعه", 3721, "سند صرف نقدي", d],
    ]
    hb = ["مدين", "دائن", "التاريخ", "البيان"]
    rb = [
        [85000, None, d, "مناولة: امين صالح المرهبي/س.ص 3691"],
        [None, 209520, d, "مناولة: بشير عتيق الحزوره//الرقم العام 5570371718"],
        [25000, None, d, "حوالة رقم 1007 باسم جلال قاسم المريسي من الادارة"],
        [25000, None, d, "حوالة رقم 3697 باسم صدام حسين الاديمي من الادارة"],   # رقم الصرافة يصادف رقم سند آخر
        [600, None, d, "خدمة حوالة باسم صديق السلامي"],
    ]
    A, B, *_ = _recs(tmp_path, ha, ra, hb, rb)
    assert A[0].roles.get("ref") == "رقم المستند"
    res = reconcile(A, B)
    pair = {r.a.row: (r.b.row, r.status) for r in res.rows if r.a and r.b}
    assert pair[A[0].row][1] == STATUS_MATCHED and pair[A[1].row][1] == STATUS_MATCHED
    # رقم 3697 يصادف سند «جلال» ورقم حوالة «صدام» في الصرافة: يجب ألا يُربطا ببعض
    assert pair[A[2].row][0] == B[2].row
    assert pair[A[3].row][0] == B[3].row
    assert res.counts[STATUS_MISSING_A] == 1 and res.counts[STATUS_MISSING_B] == 1
