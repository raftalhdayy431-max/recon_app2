"""تصدير تقرير Excel (RTL) بأوراق اختيارية: كامل أو الأوراق التي يحدّدها المستخدم فقط."""
from __future__ import annotations

import math
from datetime import datetime
from decimal import Decimal
from typing import Iterable, Optional

import xlsxwriter

from . import APP_NAME
from .analytics import _fmt_cur, analyze
from .engine import (ALL_STATUSES, BRIDGE_ORDER, SRC_A, SRC_B, STATUS_DIFF, STATUS_DUP,
                     STATUS_MATCHED, STATUS_MISSING_A, STATUS_MISSING_B, STATUS_POSSIBLE,
                     MatchRow, ReconResult)
from .loader import RawTable, SkippedRow
from .normalize import q2

# (المفتاح، اسم الورقة، الوصف) — بالترتيب الذي تظهر به في الملف
SHEET_CATALOG = [
    ("summary", "الملخص", "لوحة النتائج: المؤشرات والتوزيع والفروقات المالية وأهم العمليات"),
    ("review", "للمراجعة", "العمليات التي تحتاج قراراً مرتبة بالأولوية، مع خانات القرار والملاحظة"),
    ("matched", "المتطابقات", "كل العمليات المتطابقة وطريقة اعتمادها (تلقائي أو اعتماد المراجع)"),
    ("diff", "جدول الفروقات", "كل فرق مالي وفئته (مجموعه يساوي فرق إجمالي الملفين)"),
    ("match", "جدول المطابقة", "كل العمليات بجميع حالاتها: ملف المحاسب جنباً إلى جنب مع ملف الصرافة"),
    ("analysis", "التحليل", "ملاحظات ذكية، توزيع الفروق، الأطراف الأكثر مشاكل، والمحور الزمني"),
    ("settings", "الإعدادات والملاحظات", "الملفات والقواعد والتسامحات المستخدمة وتنبيهات التشغيل"),
    ("src_a", "بيانات المحاسب", "ملف المحاسب الأصلي كما هو مع نتيجة مطابقة كل صف"),
    ("src_b", "بيانات الصرافة", "ملف الصرافة الأصلي كما هو مع نتيجة مطابقة كل صف"),
]
SHEET_NAME = {k: n for k, n, _ in SHEET_CATALOG}
ALL_SHEETS = [k for k, _, _ in SHEET_CATALOG]

STATUS_COLORS = {
    STATUS_MATCHED: ("#DCFCE7", "#166534"),
    STATUS_DIFF: ("#FEF3C7", "#92400E"),
    STATUS_POSSIBLE: ("#DBEAFE", "#1E40AF"),
    STATUS_MISSING_B: ("#FEE2E2", "#991B1B"),
    STATUS_MISSING_A: ("#FEE2E2", "#991B1B"),
    STATUS_DUP: ("#EDE9FE", "#5B21B6"),
}
NAVY = "#0F2A43"
TEAL = "#0F766E"
TEAL_L = "#CCFBF1"
INK = "#1F2937"
MUTED = "#6B7280"
LINE = "#E2E8F0"
ZEBRA = "#F8FAFC"
ACCENT = TEAL
GRID = LINE


class _Null:
    """ورقة غير مطلوبة في التصدير: تبتلع كل الاستدعاءات."""
    def __getattr__(self, name):
        return lambda *a, **k: None

    def __bool__(self):
        return False


def _f(d) -> Optional[float]:
    if d is None:
        return None
    return float(q2(Decimal(d)))


def _lab(r: MatchRow, n: int = 70) -> str:
    """معرّف مقروء للعملية: المرجع + مقتطف من البيان."""
    rec = r.a or r.b
    d = " ".join((rec.desc or "").split()) if rec else ""
    d = d if len(d) <= n else d[:n - 1] + "…"
    ref = r.ref
    return f"{ref} · {d}" if ref and d else (ref or d or r.label)


def _est_height(text, width_chars: float, base: float = 15.0, cap: float = 78.0) -> float:
    if not text:
        return base + 3
    lines = 0
    for seg in str(text).split("\n"):
        lines += max(1, math.ceil(len(seg) / max(width_chars * 1.1, 1)))
    return max(base + 3, min(cap, base * lines + 3))


def export_report(path: str, res: ReconResult, table_a: Optional[RawTable] = None,
                  table_b: Optional[RawTable] = None, skipped: Optional[list[SkippedRow]] = None,
                  file_names: tuple[str, str] = ("", ""), sheets: Optional[Iterable[str]] = None) -> str:
    """sheets: مفاتيح الأوراق المطلوبة (انظر SHEET_CATALOG)، None = التقرير كاملاً."""
    skipped = skipped or []
    sel = set(ALL_SHEETS if sheets is None else sheets) & set(ALL_SHEETS)
    if not sel:
        raise ValueError("اختر ورقة واحدة على الأقل للتصدير.")
    if table_a is None:
        sel.discard("src_a")
    if table_b is None:
        sel.discard("src_b")

    wb = xlsxwriter.Workbook(path, {"default_date_format": "yyyy-mm-dd"})
    wb.set_properties({"title": "تقرير المطابقة والتسوية المالية", "author": APP_NAME,
                       "comments": "تم إنشاؤه آلياً"})
    base = {"font_name": "Segoe UI", "font_size": 10, "valign": "vcenter", "reading_order": 2, "font_color": INK}
    fmt = lambda **kw: wb.add_format({**base, **kw})
    bd = dict(border=1, border_color=LINE)
    F = {
        "title": fmt(bold=True, font_size=18, font_color=NAVY),
        "sub": fmt(font_size=10, font_color=MUTED),
        "h": fmt(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", text_wrap=True, **bd),
        "h2": fmt(bold=True, font_color="#FFFFFF", bg_color=TEAL, align="center", text_wrap=True, **bd),
        "sec": fmt(bold=True, font_size=12, font_color=NAVY, bottom=2, bottom_color=TEAL),
        "t": fmt(**bd),
        "tw": fmt(text_wrap=True, valign="top", **bd),
        "tc": fmt(align="center", **bd),
        "n": fmt(num_format="#,##0.00;[Red]-#,##0.00", **bd),
        "nb": fmt(num_format="#,##0.00;[Red]-#,##0.00", bold=True, **bd),
        "int": fmt(num_format="#,##0", align="center", **bd),
        "intb": fmt(num_format="#,##0", align="center", bold=True, bg_color="#F1F5F9", **bd),
        "pct": fmt(num_format="0.0%", align="center", bold=True, **bd),
        "d": fmt(num_format="yyyy-mm-dd", align="center", **bd),
        "lbl": fmt(bold=True, bg_color="#F1F5F9", **bd),
        "lblw": fmt(bold=True, bg_color="#F1F5F9", text_wrap=True, valign="top", **bd),
        "ok": fmt(bold=True, font_color="#166534", bg_color="#DCFCE7", align="center", **bd),
        "bad": fmt(bold=True, font_color="#991B1B", bg_color="#FEE2E2", align="center", **bd),
        "edit": fmt(bg_color="#FFFBEB", text_wrap=True, **bd),
        "grp": fmt(bold=True, align="center", font_color=NAVY, bg_color="#E2E8F0", **bd),
        "link": fmt(font_color="#0F766E", underline=1, bold=True, **bd),
        "banner": fmt(bold=True, font_size=22, font_color="#FFFFFF", bg_color=TEAL, align="right", indent=1),
        "banner_s": fmt(font_size=10, font_color="#CCFBF1", bg_color=TEAL, align="right", indent=1),
    }
    st_fmt = {s: fmt(bold=True, align="center", bg_color=bg, font_color=fg, **bd)
              for s, (bg, fg) in STATUS_COLORS.items()}
    prio_fmt = {
        "عالية": fmt(bold=True, align="center", font_color="#991B1B", **bd),
        "متوسطة": fmt(bold=True, align="center", font_color="#92400E", **bd),
        "منخفضة": fmt(align="center", font_color="#166534", **bd),
        "—": fmt(align="center", font_color="#94A3B8", **bd),
    }
    red = wb.add_format({"font_color": "#B91C1C", "bold": True})
    green = wb.add_format({"font_color": "#15803D"})
    zebra = wb.add_format({"bg_color": ZEBRA})

    ws_of = {}
    for key, name, _d in SHEET_CATALOG:
        ws_of[key] = wb.add_worksheet(name) if key in sel else _Null()
    for key, w in ws_of.items():
        if w:
            w.right_to_left()
            w.hide_gridlines(2)
            w.set_paper(9)
    tabs = {"summary": NAVY, "review": "#DC2626", "matched": "#16A34A", "diff": "#D97706", "match": TEAL,
            "analysis": "#7C3AED", "settings": "#64748B", "src_a": "#94A3B8", "src_b": "#94A3B8"}
    for k, c in tabs.items():
        ws_of[k].set_tab_color(c)

    rows = res.rows
    q = lambda s: "'" + s + "'"
    recs_a = [r.a for r in rows if r.a]
    recs_b = [r.b for r in rows if r.b]
    ROLE_LABEL = [("ref", "المرجع"), ("ref2", "المرجع الثانوي"), ("date", "التاريخ"), ("party", "الاسم"),
                  ("desc", "البيان"), ("fee", "الرسوم")]

    def side_cols(side_name, recs, table):
        """أعمدة هذا الملف في الجداول: الموجود فعلاً فقط (ديناميكي)."""
        cols = [(f"صف {side_name}", lambda rec: rec.row, "int", 8)]
        if res.settings.report_all_columns and table is not None and table.columns:
            for c in table.columns:
                if any(rec.raw.get(c) not in (None, "") for rec in recs):
                    cols.append((str(c), (lambda rec, c=c: rec.raw.get(c)), "any", max(10, min(30, len(str(c)) + 6))))
        else:
            for role, label in ROLE_LABEL:
                has = any((getattr(rec, role) not in (None, "", 0) and getattr(rec, role) != Decimal(0)) for rec in recs)
                if not has:
                    continue
                name = next((rec.roles[role] for rec in recs if rec.roles.get(role)), label)
                kind = "date" if role == "date" else ("num" if role == "fee" else "text")
                cols.append((str(name), (lambda rec, role=role: getattr(rec, role)), kind,
                             44 if role == "desc" else (14 if role in ("date", "fee") else 16)))
        if any(rec.currency for rec in recs):
            cols.append((f"عملة {side_name}", lambda rec: rec.currency, "tc", 8))
        cols.append((f"مبلغ {side_name}", lambda rec: _f(rec.amount), "num", 15))
        return cols

    colsA = side_cols("المحاسب", recs_a, table_a)
    colsB = side_cols("الصرافة", recs_b, table_b)

    def put(ws, x, j, v, kind):
        if v is None or v == "":
            ws.write(x, j, "", F["t"])
        elif hasattr(v, "year"):
            ws.write_datetime(x, j, datetime.combine(v, datetime.min.time()) if not isinstance(v, datetime) else v, F["d"])
        elif isinstance(v, Decimal):
            ws.write_number(x, j, _f(v), F["n"])
        elif isinstance(v, bool):
            ws.write(x, j, str(v), F["t"])
        elif isinstance(v, (int, float)):
            ws.write_number(x, j, v, F["int"] if kind == "int" else (F["n"] if isinstance(v, float) and not float(v).is_integer() else F["t"]))
        else:
            ws.write(x, j, str(v), F["tc"] if kind == "tc" else F["tw"] if kind == "text" else F["t"])

    # ======================================================== جداول المطابقة (جدول المطابقة + المتطابقات)
    def write_pairs(ws, data, title, subtitle, mode):
        head = ([("#", 6), ("الحالة", 17), ("الأولوية", 9), ("طريقة المطابقة", 16), ("الثقة %", 8)] if mode == "all"
                else [("#", 6), ("طريقة المطابقة", 18), ("الثقة %", 8), ("طريقة الاعتماد", 15)])
        tail = ([("الفرق (محاسب − صرافة)", 15), ("التفسير", 52), ("الإجراء المقترح", 40), ("قرار المراجع", 13), ("ملاحظات المراجع", 28)]
                if mode == "all" else [("الفرق (محاسب − صرافة)", 15), ("التفسير", 52), ("ملاحظات المراجع", 28)])
        headers = [h for h, _ in head] + [c[0] for c in colsA] + [c[0] for c in colsB] + [h for h, _ in tail]
        widths = [w for _, w in head] + [c[3] for c in colsA] + [c[3] for c in colsB] + [w for _, w in tail]
        nA, nB, nH = len(colsA), len(colsB), len(head)
        cA0, cB0 = nH, nH + nA
        cD = cB0 + nB
        ws.write(0, 0, title, F["title"])
        ws.write(1, 0, subtitle, F["sub"])
        ws.set_row(0, 28)
        HR = 3
        ws.merge_range(HR - 1, cA0, HR - 1, cB0 - 1, "ملف المحاسب", F["grp"])
        ws.merge_range(HR - 1, cB0, HR - 1, cD - 1, "ملف الصرافة", F["grp"])
        for j, h in enumerate(headers):
            ws.write(HR, j, h, F["h"])
            ws.set_column(j, j, widths[j])
        ws.set_row(HR, 34)
        desc_w = [(cA0 + k, c[3]) for k, c in enumerate(colsA) if c[0] not in () and c[3] >= 40] + \
                 [(cB0 + k, c[3]) for k, c in enumerate(colsB) if c[3] >= 40]
        expl_col = cD + 1
        for i, r in enumerate(data):
            x = HR + 1 + i
            ws.write_number(x, 0, r.id, F["int"])
            if mode == "all":
                ws.write(x, 1, r.status, st_fmt[r.status])
                ws.write(x, 2, r.priority, prio_fmt.get(r.priority, F["tc"]))
                ws.write(x, 3, r.method, F["tc"])
                ws.write_number(x, 4, r.confidence, F["int"])
            else:
                ws.write(x, 1, r.method, F["tc"])
                ws.write_number(x, 2, r.confidence, F["int"])
                man = r.decision == "معتمد"
                ws.write(x, 3, "اعتماد المراجع" if man else "تلقائي", F["ok"] if man else F["tc"])
            lines = []
            for base_c, cols, rec in ((cA0, colsA, r.a), (cB0, colsB, r.b)):
                for k, (_h, fn, kind, w) in enumerate(cols):
                    v = fn(rec) if rec else ""
                    if kind == "num" and v not in ("", None):
                        ws.write_number(x, base_c + k, v, F["nb"])
                    else:
                        put(ws, x, base_c + k, v, kind)
                    if kind == "text" and w >= 40 and v:
                        lines.append(_est_height(v, w))
            ws.write_number(x, cD, _f(r.diff), F["nb"])
            ws.write(x, expl_col, r.reason_text, F["tw"])
            lines.append(_est_height(r.reason_text, 52))
            if mode == "all":
                ws.write(x, expl_col + 1, r.action, F["tw"])
                ws.write(x, expl_col + 2, r.decision, F["edit"])
                ws.write(x, expl_col + 3, r.note, F["edit"])
                lines.append(_est_height(r.action, 40))
            else:
                ws.write(x, expl_col + 1, r.note, F["edit"])
            ws.set_row(x, max(lines))
        last = HR + len(data)
        if data:
            ws.autofilter(HR, 0, last, len(headers) - 1)
            ws.conditional_format(HR + 1, cD, last, cD, {"type": "cell", "criteria": "!=", "value": 0, "format": red})
            ws.conditional_format(HR + 1, cD, last, cD, {"type": "cell", "criteria": "==", "value": 0, "format": green})
            conf_col = 4 if mode == "all" else 2
            ws.conditional_format(HR + 1, conf_col, last, conf_col,
                                  {"type": "data_bar", "bar_color": "#5EEAD4", "bar_solid": True, "min_type": "num",
                                   "min_value": 0, "max_type": "num", "max_value": 100})
            if mode == "all":
                ws.data_validation(HR + 1, expl_col + 2, last, expl_col + 2,
                                   {"validate": "list", "source": ["معتمد", "مرفوض", "قيد المراجعة"]})
        ws.freeze_panes(HR + 1, 2)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        ws.repeat_rows(HR)
        return HR, last

    HR_M, last_m = write_pairs(
        ws_of["match"], rows, "جدول المطابقة التفصيلي",
        "كل عملية مع مصدرها ونتيجة مطابقتها — استخدم أزرار التصفية في العناوين للفرز والبحث.", "all")
    ok_rows = [r for r in rows if r.status == STATUS_MATCHED]
    n_manual = sum(1 for r in ok_rows if r.decision == "معتمد")
    write_pairs(ws_of["matched"], ok_rows, "العمليات المتطابقة",
                f"{len(ok_rows):,} عملية متطابقة، منها {n_manual:,} اعتمدها المراجع يدوياً (عمود «طريقة الاعتماد»).", "ok")

    # ======================================================== جدول الفروقات
    ws_d = ws_of["diff"]
    ws_d.write(0, 0, "جدول الفروقات المالية", F["title"])
    ws_d.write(1, 0, "كل فرق مالي مع فئته. مجموع هذا الجدول لكل عملة = إجمالي المحاسب − إجمالي الصرافة.", F["sub"])
    ws_d.set_row(0, 28)
    dh = ["#", "الحالة", "المرجع / البيان", "العملة", "مبلغ المحاسب", "مبلغ الصرافة", "قيمة الفرق", "فئة الفرق", "التفسير"]
    dw = [6, 18, 34, 9, 15, 15, 15, 38, 70]
    DHR = 3
    for j, h in enumerate(dh):
        ws_d.write(DHR, j, h, F["h"])
        ws_d.set_column(j, j, dw[j])
    ws_d.set_row(DHR, 30)
    x = DHR + 1
    for r in rows:
        for cur, cat, amt in r.contribs:
            if amt == 0:
                continue
            ws_d.write_number(x, 0, r.id, F["int"])
            ws_d.write(x, 1, r.status, st_fmt[r.status])
            ws_d.write(x, 2, _lab(r, 60), F["tw"])
            ws_d.write(x, 3, cur, F["tc"])
            ws_d.write(x, 4, _f(r.amt_a) if r.amt_a is not None else "", F["n"])
            ws_d.write(x, 5, _f(r.amt_b) if r.amt_b is not None else "", F["n"])
            ws_d.write_number(x, 6, _f(amt), F["nb"])
            ws_d.write(x, 7, cat, F["t"])
            ws_d.write(x, 8, r.reason_text, F["tw"])
            ws_d.set_row(x, _est_height(r.reason_text, 70))
            x += 1
    if x > DHR + 1:
        ws_d.autofilter(DHR, 0, x - 1, len(dh) - 1)
        ws_d.conditional_format(DHR + 1, 6, x - 1, 6, {"type": "cell", "criteria": ">", "value": 0, "format": green})
        ws_d.conditional_format(DHR + 1, 6, x - 1, 6, {"type": "cell", "criteria": "<", "value": 0, "format": red})
    ws_d.freeze_panes(DHR + 1, 0)
    ws_d.set_landscape()
    ws_d.fit_to_pages(1, 0)

    # ======================================================== للمراجعة
    ws_r = ws_of["review"]
    ws_r.write(0, 0, "العمليات التي تحتاج إلى مراجعة", F["title"])
    ws_r.write(1, 0, "مرتبة حسب الأولوية. سجّل قرارك وملاحظتك في الأعمدة الصفراء.", F["sub"])
    ws_r.set_row(0, 28)
    rh = ["#", "الأولوية", "الحالة", "المصدر", "المرجع / البيان", "العملة", "قيمة الفرق", "ما المشكلة؟", "الإجراء المقترح",
          "قرار المراجع", "ملاحظات المراجع"]
    rw = [6, 10, 18, 10, 34, 9, 15, 62, 52, 14, 30]
    RHR = 3
    for j, h in enumerate(rh):
        ws_r.write(RHR, j, h, F["h"])
        ws_r.set_column(j, j, rw[j])
    ws_r.set_row(RHR, 30)
    prio_rank = {"عالية": 0, "متوسطة": 1, "منخفضة": 2}
    review = [r for r in rows if r.status != STATUS_MATCHED or r.priority == "منخفضة"]
    review.sort(key=lambda r: (prio_rank.get(r.priority, 3), -abs(r.diff)))
    row_in_review = {}
    for i, r in enumerate(review):
        x = RHR + 1 + i
        row_in_review[r.id] = x
        ws_r.write_number(x, 0, r.id, F["int"])
        ws_r.write(x, 1, r.priority, prio_fmt.get(r.priority, F["tc"]))
        ws_r.write(x, 2, r.status, st_fmt[r.status])
        ws_r.write(x, 3, r.source, F["tc"])
        ws_r.write(x, 4, _lab(r), F["tw"])
        ws_r.write(x, 5, r.currency, F["tc"])
        ws_r.write_number(x, 6, _f(r.diff), F["nb"])
        ws_r.write(x, 7, r.reason_text, F["tw"])
        ws_r.write(x, 8, r.action, F["tw"])
        ws_r.write(x, 9, r.decision, F["edit"])
        ws_r.write(x, 10, r.note, F["edit"])
        ws_r.set_row(x, max(_est_height(r.reason_text, 62), _est_height(r.action, 52), _est_height(_lab(r), 34)))
    if review:
        ws_r.autofilter(RHR, 0, RHR + len(review), len(rh) - 1)
        ws_r.data_validation(RHR + 1, 9, RHR + len(review), 9,
                             {"validate": "list", "source": ["معتمد", "مرفوض", "قيد المراجعة"]})
        ws_r.conditional_format(RHR + 1, 6, RHR + len(review), 6, {"type": "cell", "criteria": ">", "value": 0, "format": green})
        ws_r.conditional_format(RHR + 1, 6, RHR + len(review), 6, {"type": "cell", "criteria": "<", "value": 0, "format": red})
    ws_r.freeze_panes(RHR + 1, 0)
    ws_r.set_landscape()
    ws_r.fit_to_pages(1, 0)

    # ======================================================== التحليل
    _write_analysis(wb, ws_of["analysis"], F, analyze(res) if "analysis" in sel else None)

    # ======================================================== بيانات المصدر مع النتيجة
    by_a = {r.a.row: r for r in rows if r.a}
    by_b = {r.b.row: r for r in rows if r.b}
    for r in rows:
        for m in r.members_a:
            by_a[m.row] = r
        for m in r.members_b:
            by_b[m.row] = r

    def write_source(ws, table: RawTable, lookup, skipped_rows, side_title):
        ws.write(0, 0, side_title, F["title"])
        ws.write(1, 0, f"الملف: {table.name} — ورقة: {table.sheet}. العمودان الأولان يوضحان نتيجة مطابقة كل صف.", F["sub"])
        ws.set_row(0, 28)
        hh = ["نتيجة المطابقة", "# في الجدول"] + table.columns
        HR2 = 3
        for j, h in enumerate(hh):
            ws.write(HR2, j, h, F["h2"] if j < 2 else F["h"])
            ws.set_column(j, j, 20 if j < 2 else max(12, min(34, len(str(h)) + 6)))
        ws.set_row(HR2, 30)
        sk = {s.row: s for s in skipped_rows}
        for i, (xl_row, vals) in enumerate(table.rows):
            x = HR2 + 1 + i
            r = lookup.get(xl_row)
            if r:
                ws.write(x, 0, r.status, st_fmt[r.status])
                ws.write_number(x, 1, r.id, F["int"])
            else:
                ws.write(x, 0, "مُهمل: " + sk[xl_row].reason if xl_row in sk else "—", F["t"])
                ws.write(x, 1, "", F["t"])
            for j, v in enumerate(vals):
                if hasattr(v, "year"):
                    ws.write_datetime(x, j + 2, datetime.combine(v, datetime.min.time())
                                      if not isinstance(v, datetime) else v, F["d"])
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    ws.write_number(x, j + 2, v, F["n"] if isinstance(v, float) else F["t"])
                else:
                    ws.write(x, j + 2, "" if v is None else str(v), F["t"])
        ws.autofilter(HR2, 0, HR2 + len(table.rows), len(hh) - 1)
        ws.freeze_panes(HR2 + 1, 2)
        ws.conditional_format(HR2 + 1, 0, HR2 + len(table.rows), len(hh) - 1,
                              {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": zebra})

    if table_a:
        write_source(ws_of["src_a"], table_a, by_a, [s for s in skipped if s.side == "A"],
                     "بيانات ملف المحاسب مع نتيجة المطابقة")
    if table_b:
        write_source(ws_of["src_b"], table_b, by_b, [s for s in skipped if s.side == "B"],
                     "بيانات ملف الصرافة مع نتيجة المطابقة")

    # ======================================================== الإعدادات والملاحظات
    ws = ws_of["settings"]
    ws.set_column(0, 0, 3)
    ws.set_column(1, 1, 40)
    ws.set_column(2, 2, 90)
    ws.write(0, 1, "الإعدادات والملاحظات", F["title"])
    ws.write(1, 1, "ما الذي اعتمده البرنامج في هذه المطابقة، وأي تنبيهات رافقت التشغيل.", F["sub"])
    ws.set_row(0, 28)
    S = res.settings
    cfg = [("ملف المحاسب", file_names[0] or (table_a.name if table_a else "")),
           ("ملف الصرافة", file_names[1] or (table_b.name if table_b else "")),
           ("فترة المحاسب", _rng(res.range_a)), ("فترة الصرافة", _rng(res.range_b)),
           ("اتجاه الإشارات", res.rel_note),
           ("تسامح المبلغ", f"{S.amount_tol}"), ("تسامح التاريخ (بالمرجع)", f"{S.date_tol_days} يوم"),
           ("نافذة التاريخ للمطابقة الذكية", f"{S.fuzzy_days} يوم"),
           ("أعلى نسبة فرق للمطابقة المحتملة", f"{S.possible_pct * 100:.1f}%"),
           ("المطابقة الذكية", "مفعّلة" if S.use_fuzzy else "معطّلة"),
           ("صرامة العملة", "مفعّلة" if S.strict_currency else "معطّلة")]
    for k, kr in enumerate(getattr(res, "key_rules", []) or [], 1):
        cfg.append((f"مفتاح المطابقة {k}: {kr.name}", f"المحاسب: {'، '.join(kr.a)}   ⟷   الصرافة: {'، '.join(kr.b)}"))
    xr = 3
    ws.write(xr, 1, "الإعدادات المستخدمة", F["sec"])
    for i, (k, v) in enumerate(cfg, 1):
        ws.write(xr + i, 1, k, F["lbl"])
        ws.write(xr + i, 2, str(v), F["tw"])
        ws.set_row(xr + i, _est_height(v, 90))
    xr += len(cfg) + 2
    notes = list(res.warnings)
    if skipped:
        notes.append(f"تم إهمال {len(skipped)} صفاً (إجماليات/بلا مبلغ) ولم تدخل في المطابقة — تظهر في ورقتي البيانات.")
    ws.write(xr, 1, "ملاحظات التشغيل", F["sec"])
    if notes:
        for i, w in enumerate(notes, 1):
            ws.write(xr + i, 1, f"ملاحظة {i}", F["lbl"])
            ws.write(xr + i, 2, w, F["tw"])
            ws.set_row(xr + i, _est_height(w, 90))
    else:
        ws.write(xr + 1, 1, "—", F["lbl"])
        ws.write(xr + 1, 2, "لا توجد ملاحظات.", F["t"])

    # ======================================================== الملخص (لوحة)
    ws = ws_of["summary"]
    ws.set_column(0, 0, 2)
    ws.set_column(1, 12, 14)
    ws.set_row(0, 36)
    ws.set_row(1, 20)
    ws.merge_range(0, 1, 0, 12, "تقرير المطابقة والتسوية المالية", F["banner"])
    names = " ⟷ ".join(n for n in ((file_names[0] or (table_a.name if table_a else "")),
                                   (file_names[1] or (table_b.name if table_b else ""))) if n)
    ws.merge_range(1, 1, 1, 12, f"{names}   ·   {datetime.now():%Y-%m-%d %H:%M}   ·   {APP_NAME}", F["banner_s"])

    c = res.counts
    denom = max(res.count_a, res.count_b, 1)
    need = sum(1 for r in rows if r.status != STATUS_MATCHED)
    high = sum(1 for r in rows if r.priority == "عالية")
    exp = {}
    for r in rows:
        if r.status == STATUS_MATCHED:
            continue
        for cur, _cat, amt in (r.contribs or [(r.currency, "", r.diff)]):
            exp[cur or ""] = exp.get(cur or "", Decimal(0)) + abs(amt)
    main_cur = max(exp, key=lambda k: exp[k]) if exp else ""
    tiles = [
        ("نسبة التطابق", f"{res.match_rate * 100:.1f}%", f"{c.get(STATUS_MATCHED, 0):,} عملية متطابقة", "#16A34A", "#F0FDF4"),
        ("تحتاج مراجعة", f"{need:,}", f"منها {high:,} بأولوية عالية", "#DC2626", "#FEF2F2"),
        ("القيمة غير المطابقة", f"{_f(exp.get(main_cur, Decimal(0))):,.2f}" if exp else "0", f"بعملة {main_cur or '—'}", "#D97706", "#FFFBEB"),
        ("عدد العمليات", f"{res.count_a:,} | {res.count_b:,}", "محاسب | صرافة", "#2563EB", "#EFF6FF"),
    ]
    ws.set_row(3, 18)
    ws.set_row(4, 34)
    ws.set_row(5, 18)
    for i, (lbl, val, sub, col, bg) in enumerate(tiles):
        c0 = 1 + i * 3
        ws.merge_range(3, c0, 3, c0 + 2, lbl, fmt(bold=True, font_color=MUTED, bg_color=bg, align="center", top=2, top_color=col))
        ws.merge_range(4, c0, 4, c0 + 2, val, fmt(bold=True, font_size=22, font_color=col, bg_color=bg, align="center"))
        ws.merge_range(5, c0, 5, c0 + 2, sub, fmt(font_size=9, font_color=MUTED, bg_color=bg, align="center"))

    # نتائج المطابقة
    r1 = 7
    ws.merge_range(r1, 1, r1, 6, "نتائج المطابقة", F["sec"])
    ws.merge_range(r1 + 1, 1, r1 + 1, 2, "الحالة", F["h"])
    ws.write(r1 + 1, 3, "العدد", F["h"])
    ws.write(r1 + 1, 4, "النسبة", F["h"])
    status_rng = f"{q(SHEET_NAME['match'])}!$B${HR_M + 2}:$B${max(last_m + 1, HR_M + 2)}" if "match" in sel else None
    for i, s in enumerate(ALL_STATUSES):
        rr = r1 + 2 + i
        ws.merge_range(rr, 1, rr, 2, s, st_fmt[s])
        if status_rng:
            ws.write_formula(rr, 3, f"=COUNTIF({status_rng},B{rr + 1})", F["intb"], c.get(s, 0))
        else:
            ws.write_number(rr, 3, c.get(s, 0), F["intb"])
        ws.write_formula(rr, 4, f"=D{rr + 1}/{denom}", F["pct"], c.get(s, 0) / denom)
    first_s, last_s = r1 + 2, r1 + 1 + len(ALL_STATUSES)
    ws.conditional_format(first_s, 3, last_s, 3, {"type": "data_bar", "bar_color": "#99F6E4", "bar_solid": True,
                                                  "min_type": "num", "min_value": 0})
    rr = r1 + 2 + len(ALL_STATUSES)
    ws.merge_range(rr, 1, rr, 2, "الإجمالي", F["lbl"])
    ws.write_formula(rr, 3, f"=SUM(D{first_s + 1}:D{last_s + 1})", F["intb"], len(rows))
    chart_first = first_s

    # الفروقات المالية والجسر
    cur_list = res.currencies or [""]
    r2 = rr + 3
    ws.merge_range(r2, 1, r2, 6, "الفروقات المالية (لكل عملة)", F["sec"])
    ws.merge_range(r2 + 1, 1, r2 + 1, 3, "البند", F["h"])
    for j, cu in enumerate(cur_list[:3]):
        ws.write(r2 + 1, 4 + j, cu or "غير محدد", F["h"])
    lines = [("إجمالي ملف المحاسب", lambda cu: res.total_a.get(cu, Decimal(0)), False),
             ("إجمالي ملف الصرافة (بعد مواءمة الإشارة)", lambda cu: res.total_b.get(cu, Decimal(0)), False),
             ("الفرق الصافي (محاسب − صرافة)", lambda cu: res.net_diff(cu), True)]
    rix = r2 + 2
    for name, fn, bold in lines:
        ws.merge_range(rix, 1, rix, 3, name, F["lbl"])
        for j, cu in enumerate(cur_list[:3]):
            ws.write_number(rix, 4 + j, _f(fn(cu)), F["nb"] if bold else F["n"])
        rix += 1
    rix += 1
    ws.merge_range(rix, 1, rix, 3, "تحليل الفرق الصافي حسب السبب", F["h2"])
    for j, cu in enumerate(cur_list[:3]):
        ws.write(rix, 4 + j, cu or "غير محدد", F["h2"])
    rix += 1
    b_first = rix
    for cat in BRIDGE_ORDER:
        if not any(cat in res.bridge.get(cu, {}) for cu in cur_list):
            continue
        ws.merge_range(rix, 1, rix, 3, cat, F["t"])
        for j, cu in enumerate(cur_list[:3]):
            ws.write_number(rix, 4 + j, _f(res.bridge.get(cu, {}).get(cat, Decimal(0))), F["n"])
        rix += 1
    b_last = rix - 1
    ws.merge_range(rix, 1, rix, 3, "مجموع الأسباب", F["lbl"])
    for j, cu in enumerate(cur_list[:3]):
        col = chr(ord("E") + j)
        ws.write_formula(rix, 4 + j, f"=SUM({col}{b_first + 1}:{col}{b_last + 1})" if b_last >= b_first else "=0", F["nb"],
                         _f(sum(res.bridge.get(cu, {}).values(), Decimal(0))))
    rix += 1
    ws.merge_range(rix, 1, rix, 3, "التحقق الحسابي (يجب أن يساوي صفراً)", F["lbl"])
    for j, cu in enumerate(cur_list[:3]):
        okv = res.integrity.get(cu, Decimal(0)) == 0
        ws.write(rix, 4 + j, "✔ مغلق" if okv else f"✖ فرق {res.integrity.get(cu)}", F["ok"] if okv else F["bad"])
    rix += 2

    # أهم العمليات للمراجعة
    top = review[:10]
    if top:
        ws.merge_range(rix, 1, rix, 12, "أهم العمليات للمراجعة (الأعلى أولوية ثم الأكبر قيمة)", F["sec"])
        rix += 1
        for c0, c1, h in ((1, 1, "#"), (2, 2, "الأولوية"), (3, 4, "الحالة"), (5, 7, "المرجع / البيان"),
                          (8, 8, "العملة"), (9, 9, "الفرق"), (10, 12, "الإجراء المقترح")):
            (ws.merge_range(rix, c0, rix, c1, h, F["h"]) if c1 > c0 else ws.write(rix, c0, h, F["h"]))
        rix += 1
        for r in top:
            ws.set_row(rix, max(_est_height(r.action, 40), _est_height(_lab(r, 90), 42)))
            link = f"internal:{q(SHEET_NAME['review'])}!A{row_in_review.get(r.id, 0) + 1}" if "review" in sel else None
            (ws.write_url(rix, 1, link, F["link"], str(r.id)) if link else ws.write_number(rix, 1, r.id, F["int"]))
            ws.write(rix, 2, r.priority, prio_fmt.get(r.priority, F["tc"]))
            ws.merge_range(rix, 3, rix, 4, r.status, st_fmt[r.status])
            ws.merge_range(rix, 5, rix, 7, _lab(r, 90), F["tw"])
            ws.write(rix, 8, r.currency, F["tc"])
            ws.write_number(rix, 9, _f(r.diff), F["nb"])
            ws.merge_range(rix, 10, rix, 12, r.action, F["tw"])
            rix += 1
        rix += 1

    # فهرس الأوراق
    ws.merge_range(rix, 1, rix, 12, "ماذا يحتوي هذا التقرير؟", F["sec"])
    rix += 1
    for key, name, desc in SHEET_CATALOG:
        if key not in sel or key == "summary":
            continue
        ws.merge_range(rix, 1, rix, 3, "", F["link"])
        ws.write_url(rix, 1, f"internal:{q(name)}!A1", F["link"], name)
        ws.merge_range(rix, 4, rix, 12, desc, F["t"])
        rix += 1

    # رسوم
    ch = wb.add_chart({"type": "doughnut"})
    ch.add_series({
        "name": "نتائج المطابقة",
        "categories": [SHEET_NAME["summary"], chart_first, 1, chart_first + len(ALL_STATUSES) - 1, 1],
        "values": [SHEET_NAME["summary"], chart_first, 3, chart_first + len(ALL_STATUSES) - 1, 3],
        "points": [{"fill": {"color": col}} for col in ("#16A34A", "#F59E0B", "#3B82F6", "#EF4444", "#B91C1C", "#8B5CF6")],
        "data_labels": {"value": True},
    })
    ch.set_title({"name": "توزيع نتائج المطابقة", "name_font": {"size": 11}})
    ch.set_legend({"position": "bottom"})
    ch.set_size({"width": 400, "height": 250})
    ws.insert_chart(r1, 7, ch, {"x_offset": 8})
    if b_last >= b_first:
        bc = wb.add_chart({"type": "bar"})
        for j, cu in enumerate(cur_list[:3]):
            bc.add_series({"name": cu or "غير محدد",
                           "categories": [SHEET_NAME["summary"], b_first, 1, b_last, 1],
                           "values": [SHEET_NAME["summary"], b_first, 4 + j, b_last, 4 + j]})
        bc.set_title({"name": "أسباب الفرق الصافي", "name_font": {"size": 11}})
        bc.set_legend({"position": "bottom"})
        bc.set_size({"width": 400, "height": 290})
        ws.insert_chart(r1 + 13, 7, bc, {"x_offset": 8})
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.activate()
    wb.close()
    return path


def _write_analysis(wb, ws, F, an) -> None:
    """ورقة «التحليل»: ملاحظات، توزيع الفروق، الأطراف، المحور الزمني (مع رسوم بيانية)."""
    if an is None or not ws:
        return
    ws.set_column(0, 0, 4)
    ws.set_column(1, 1, 52)
    ws.set_column(2, 5, 18)
    ws.write(0, 1, "تحليل نتائج المطابقة", F["title"])
    ws.write(1, 1, f"{an.issues_total:,} عملية تحتاج متابعة من أصل {an.rows_total:,}", F["sub"])
    x = 3
    ws.write(x, 1, "ملاحظات ذكية", F["sec"])
    x += 1
    tag = {"high": "مهم", "mid": "انتبه", "info": "معلومة", "good": "جيد"}
    for f in an.findings:
        ws.write(x, 1, f"[{tag.get(f.level, '')}] {f.title}", F["lbl"])
        ws.merge_range(x, 2, x, 5, f.detail, F["tw"])
        ws.set_row(x, 44)
        x += 1
    x += 1
    ws.write(x, 1, "القيمة غير المطابقة لكل عملة", F["sec"])
    x += 1
    for j, h in enumerate(["العملة", "القيمة المطلقة", "الصافي (محاسب − صرافة)", "عمليات"]):
        ws.write(x, 1 + j, h, F["h"])
    x += 1
    for cur, e in sorted(an.exposure.items(), key=lambda kv: -kv[1]["abs"]):
        ws.write(x, 1, cur or "—", F["tc"])
        ws.write(x, 2, float(e["abs"]), F["n"])
        ws.write(x, 3, float(e["net"]), F["n"])
        ws.write(x, 4, e["n"], F["int"])
        x += 1
    x += 1
    ws.write(x, 1, "توزيع الفروق حسب السبب", F["sec"])
    x += 1
    for j, h in enumerate(["السبب", "العملة", "عدد العمليات", "القيمة المطلقة"]):
        ws.write(x, 1 + j, h, F["h"])
    x += 1
    c0 = x
    for c in an.categories:
        ws.write(x, 1, c.category, F["t"])
        ws.write(x, 2, c.currency or "—", F["tc"])
        ws.write(x, 3, c.count, F["int"])
        ws.write(x, 4, float(c.amount), F["n"])
        x += 1
    if an.categories:
        ch = wb.add_chart({"type": "bar"})
        ch.add_series({"categories": ["التحليل", c0, 1, x - 1, 1], "values": ["التحليل", c0, 4, x - 1, 4],
                       "fill": {"color": "#0D9488"}, "name": "القيمة"})
        ch.set_title({"name": "توزيع الفروق حسب السبب"})
        ch.set_legend({"none": True})
        ch.set_size({"width": 620, "height": max(260, 26 * len(an.categories) + 90)})
        ws.insert_chart(x + 1, 1, ch)
        x += int(max(260, 26 * len(an.categories) + 90) / 20) + 3
    if an.parties:
        ws.write(x, 1, "الأطراف الأكثر مشاكل", F["sec"])
        x += 1
        for j, h in enumerate(["الاسم", "عمليات غير مطابقة", "كل عملياته", "النسبة", "القيمة"]):
            ws.write(x, 1 + j, h, F["h"])
        x += 1
        for p in an.parties:
            ws.write(x, 1, p.party, F["t"])
            ws.write(x, 2, p.issues, F["int"])
            ws.write(x, 3, p.total, F["int"])
            ws.write(x, 4, p.rate, F["pct"])
            ws.write(x, 5, _fmt_cur(p.exposure), F["t"])
            x += 1
        x += 1
    if len(an.timeline) > 1:
        ws.write(x, 1, f"المحور الزمني (لكل {an.timeline_unit})", F["sec"])
        x += 1
        for j, h in enumerate(["الفترة", "كل العمليات", "غير المطابقة"]):
            ws.write(x, 1 + j, h, F["h"])
        x += 1
        t0 = x
        for lbl, tot, bad in an.timeline:
            ws.write(x, 1, lbl, F["tc"])
            ws.write(x, 2, tot, F["int"])
            ws.write(x, 3, bad, F["int"])
            x += 1
        ch = wb.add_chart({"type": "column"})
        ch.add_series({"categories": ["التحليل", t0, 1, x - 1, 1], "values": ["التحليل", t0, 2, x - 1, 2],
                       "fill": {"color": "#CBD5E1"}, "name": "كل العمليات"})
        ch.add_series({"categories": ["التحليل", t0, 1, x - 1, 1], "values": ["التحليل", t0, 3, x - 1, 3],
                       "fill": {"color": "#D64535"}, "name": "غير المطابقة"})
        ch.set_title({"name": "العمليات وغير المطابق منها عبر الزمن"})
        ch.set_size({"width": 720, "height": 300})
        ws.insert_chart(x + 1, 1, ch)


def _rng(rg) -> str:
    lo, hi = rg
    return f"{lo:%Y-%m-%d} → {hi:%Y-%m-%d}" if lo and hi else "غير متوفرة"
