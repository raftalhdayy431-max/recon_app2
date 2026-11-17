"""قراءة ملفات Excel/CSV، اكتشاف صف العناوين، ربط الأعمدة تلقائياً، وبناء السجلات الموحدة."""
from __future__ import annotations

import csv
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any, Optional

from .normalize import (clean_text, norm_currency, norm_ref, parse_amount,
                        parse_date)

# ---------------------------------------------------------------------------
# الحقول المنطقية وأسماء الأعمدة المحتملة (عربي/إنجليزي)
# ---------------------------------------------------------------------------
FIELDS = {
    "ref": "رقم المرجع / الحوالة",
    "ref2": "رقم مرجعي ثانوي (سند/إيصال)",
    "date": "التاريخ",
    "amount": "المبلغ",
    "debit": "مدين",
    "credit": "دائن",
    "currency": "العملة",
    "fee": "الرسوم / العمولة",
    "party": "الاسم (عميل/مستفيد)",
    "desc": "البيان",
}
REQUIRED_ONE_OF_AMOUNT = ("amount", "debit", "credit")

SYNONYMS = {
    "ref": ["رقم المستند", "المستند", "رقم المرجع", "المرجع", "رقم الحوالة", "الحوالة", "رقم العملية", "رقم المعاملة",
            "رقم الاشعار", "رقم الإشعار", "مرجع العملية", "مرجع", "رقم التحويل", "الرقم المرجعي",
            "ref", "ref no", "reference", "reference no", "reference number", "transaction id",
            "transaction no", "trx", "trx id", "txn", "txn id", "remittance", "remittance no",
            "transfer no", "order id"],
    "ref2": ["رقم السند", "السند", "رقم القيد", "القيد", "رقم الايصال", "رقم الإيصال", "الايصال",
             "رقم الفاتورة", "voucher", "voucher no", "receipt", "receipt no", "entry no", "journal"],
    "date": ["التاريخ", "تاريخ العملية", "تاريخ الحوالة", "تاريخ القيد", "تاريخ المعاملة", "تاريخ",
             "date", "value date", "trans date", "transaction date", "posting date"],
    "amount": ["المبلغ", "القيمة", "مبلغ الحوالة", "مبلغ العملية", "الصافي", "صافي المبلغ", "قيمة الحوالة",
               "amount", "value", "net amount", "total amount", "amt"],
    "debit": ["مدين", "المدين", "debit", "dr", "dr."],
    "credit": ["دائن", "الدائن", "credit", "cr", "cr."],
    "currency": ["العملة", "عملة", "رمز العملة", "currency", "ccy", "cur", "curr"],
    "fee": ["الرسوم", "رسوم", "العمولة", "عمولة", "مصاريف", "رسوم التحويل", "عمولة الحوالة",
            "fee", "fees", "commission", "charge", "charges", "comm"],
    "party": ["اسم العميل", "العميل", "المستفيد", "اسم المستفيد", "المرسل", "اسم المرسل", "الاسم",
              "الطرف", "اسم الطرف", "الحساب", "اسم الحساب", "customer", "customer name", "beneficiary",
              "sender", "name", "party", "account name", "client"],
    "desc": ["البيان", "الوصف", "التفاصيل", "ملاحظات", "الملاحظات", "شرح", "بيان العملية",
             "description", "details", "narration", "notes", "remarks", "memo"],
}

TOTAL_WORDS = ("اجمالي", "إجمالي", "المجموع", "مجموع", "الاجمالي", "الإجمالي", "total", "subtotal",
               "الرصيد", "رصيد")


@dataclass
class RawTable:
    path: str
    sheet: str
    sheets: list[str]
    header_row: int  # رقم الصف في Excel (يبدأ من 1)
    columns: list[str]
    rows: list[tuple[int, list[Any]]]  # (رقم الصف في Excel, القيم)

    @property
    def name(self) -> str:
        return Path(self.path).name

    def col_index(self, col: str) -> int:
        return self.columns.index(col)

    def preview(self, n: int = 12) -> list[list[Any]]:
        return [r[1] for r in self.rows[:n]]


@dataclass
class Record:
    side: str                 # "A" ملف المحاسب، "B" ملف الصرافة
    row: int                  # رقم الصف في الملف الأصلي
    ref: str
    ref_norm: str
    ref2: str
    date: Optional[date]
    amount: Decimal           # بإشارته كما في الملف (مدين - دائن إن وُجدت أعمدة)
    currency: str
    fee: Decimal
    party: str
    desc: str
    raw: dict = field(default_factory=dict)
    roles: dict = field(default_factory=dict)   # الدور المنطقي -> اسم العمود في الملف (المتاح فقط)
    uid: str = ""

    def __post_init__(self):
        if not self.uid:
            self.uid = f"{self.side}{self.row}"


# ---------------------------------------------------------------------------
# القراءة
# ---------------------------------------------------------------------------
def _cell_to_str_header(v) -> str:
    return "" if v is None else str(v).strip()


def _read_xlsx(path: str, sheet: Optional[str]) -> tuple[list[str], str, list[list[Any]]]:
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    names = wb.sheetnames
    if sheet is None:
        # أول ورقة تحتوي بيانات كافية
        best, best_n = names[0], -1
        for n in names:
            ws = wb[n]
            cnt = 0
            for i, _ in enumerate(ws.iter_rows(values_only=True)):
                cnt += 1
                if cnt > 5000:
                    break
            if cnt > best_n:
                best, best_n = n, cnt
        sheet = best
    ws = wb[sheet]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return names, sheet, rows


def _read_csv(path: str) -> list[list[Any]]:
    raw = Path(path).read_bytes()
    for enc in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return [list(r) for r in csv.reader(text.splitlines(), dialect)]


def _read_xls(path: str, sheet: Optional[str]):
    import pandas as pd
    xl = pd.ExcelFile(path)
    names = xl.sheet_names
    sheet = sheet or names[0]
    df = xl.parse(sheet, header=None, dtype=object)
    rows = [[None if (isinstance(x, float) and x != x) else x for x in r] for r in df.values.tolist()]
    return names, sheet, rows


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
_LATIN_RUN = re.compile(r"[A-Za-z0-9][A-Za-z0-9.,:/\-]*|[^A-Za-z0-9]+")
_ARABIC = re.compile(r"[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]")


def _pdf_cell(v):
    if v is None:
        return None
    t = unicodedata.normalize("NFKC", str(v)).replace("\n", " ").strip()
    return t or None


def _unvisual(s: str) -> str:
    """يعيد ترتيب نص عربي استُخرج بترتيب العرض (معكوس) إلى الترتيب المنطقي، مع إبقاء الأرقام."""
    if not isinstance(s, str) or not _ARABIC.search(s):
        return s
    runs = _LATIN_RUN.findall(s)
    out = []
    for r in runs:
        out.append(r if re.match(r"[A-Za-z0-9]", r) else r[::-1])
    return "".join(reversed(out))


def _read_pdf(path: str) -> list[list[Any]]:
    try:
        import pdfplumber
    except ImportError as ex:  # pragma: no cover
        raise ValueError("قراءة PDF تتطلب المكتبة pdfplumber (pip install pdfplumber).") from ex
    rows: list[list[Any]] = []
    text_lines = 0
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                tables = page.extract_tables({"vertical_strategy": "text", "horizontal_strategy": "text",
                                              "snap_tolerance": 4, "intersection_tolerance": 6})
            for t in tables:
                for r in t:
                    cells = [_pdf_cell(c) for c in r]
                    if any(c for c in cells):
                        rows.append(cells)
            text_lines += len((page.extract_text() or "").strip())
    if not rows:
        if text_lines == 0:
            raise ValueError("ملف PDF لا يحوي نصاً قابلاً للقراءة (يبدو ممسوحاً ضوئياً كصورة). "
                             "حوّله إلى Excel أو استخدم نسخة نصية من الملف.")
        raise ValueError("لم يُعثر على جدول بيانات داخل ملف PDF.")
    # بعض ملفات PDF العربية تُستخرج بترتيب معكوس: نجرّب الإصلاح ونعتمد الأفضل
    base = max((_header_score(r) for r in rows[:40]), default=0)
    fixed = [[_unvisual(c) if isinstance(c, str) else c for c in r] for r in rows]
    if max((_header_score(r) for r in fixed[:40]), default=0) > base:
        rows = fixed
    return rows


def _header_score(cells: list[Any]) -> float:
    score = 0.0
    seen = 0
    for c in cells:
        t = clean_text(c)
        if not t:
            continue
        seen += 1
        for syns in SYNONYMS.values():
            if any(t == clean_text(s) for s in syns):
                score += 1.0
                break
            if any(clean_text(s) in t for s in syns if len(clean_text(s)) > 3):
                score += 0.5
                break
    if seen < 2:
        return 0.0
    # العناوين نصوص غالباً
    text_cells = sum(1 for c in cells if isinstance(c, str) and c.strip())
    return score + 0.1 * text_cells


def load_table(path: str, sheet: Optional[str] = None, header_row: Optional[int] = None) -> RawTable:
    ext = Path(path).suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        names, sheet, rows = _read_xlsx(path, sheet)
    elif ext == ".xls":
        names, sheet, rows = _read_xls(path, sheet)
    elif ext == ".pdf":
        names, sheet, rows = ["PDF"], "PDF", _read_pdf(path)
    elif ext in (".csv", ".txt", ".tsv"):
        names, sheet, rows = ["CSV"], "CSV", _read_csv(path)
    else:
        raise ValueError(f"صيغة الملف غير مدعومة: {ext}")

    rows = [r for r in rows]
    if not any(any(c not in (None, "") for c in r) for r in rows):
        raise ValueError("الملف فارغ أو لا يحتوي بيانات.")

    if header_row is None:
        best_i, best_s = 0, -1.0
        for i, r in enumerate(rows[:40]):
            s = _header_score(r)
            if s > best_s:
                best_i, best_s = i, s
        hi = best_i
    else:
        hi = header_row - 1

    header = rows[hi]
    width = max(len(header), max((len(r) for r in rows[hi + 1:hi + 200]), default=0))
    cols: list[str] = []
    for j in range(width):
        name = _cell_to_str_header(header[j]) if j < len(header) else ""
        if not name:
            name = f"عمود {j + 1}"
        base, k = name, 2
        while name in cols:
            name = f"{base} ({k})"
            k += 1
        cols.append(name)

    body: list[tuple[int, list[Any]]] = []
    for i in range(hi + 1, len(rows)):
        r = list(rows[i]) + [None] * (width - len(rows[i]))
        r = r[:width]
        if all(c in (None, "") or (isinstance(c, str) and not c.strip()) for c in r):
            continue
        if i > hi and [_cell_to_str_header(c) for c in r[:len(header)]] == [_cell_to_str_header(c) for c in header[:width]] and any(_cell_to_str_header(c) for c in header):
            continue   # ترويسة مكررة (تتكرر في كل صفحة من ملفات PDF)
        body.append((i + 1, r))
    # أعمدة بلا عنوان وفارغة تماماً (بقايا تنسيق) تُحذف حتى لا تشوّش الربط والتقرير
    keep = [j for j, c in enumerate(cols)
            if not (re.fullmatch(r"عمود \d+", c) and all(r[1][j] in (None, "") or (isinstance(r[1][j], str) and not r[1][j].strip()) for r in body))]
    if len(keep) < len(cols):
        cols = [cols[j] for j in keep]
        body = [(i, [r[j] for j in keep]) for i, r in body]
    return RawTable(path=path, sheet=sheet, sheets=names, header_row=hi + 1, columns=cols, rows=body)


# ---------------------------------------------------------------------------
# الربط التلقائي للأعمدة
# ---------------------------------------------------------------------------
def _column_values(table: RawTable, j: int, limit: int = 300) -> list[Any]:
    return [r[1][j] for r in table.rows[:limit] if r[1][j] not in (None, "")]


def _share(vals: list[Any], pred) -> float:
    if not vals:
        return 0.0
    return sum(1 for v in vals if pred(v)) / len(vals)


def auto_map(table: RawTable) -> dict[str, Optional[str]]:
    """يُرجع {الحقل: اسم العمود} بأفضل تخمين، بالاعتماد على الاسم ثم على محتوى العمود."""
    n = len(table.columns)
    scores: dict[str, list[float]] = {f: [0.0] * n for f in FIELDS}

    for j, col in enumerate(table.columns):
        t = clean_text(col)
        vals = _column_values(table, j)
        date_share = _share(vals, lambda v: parse_date(v) is not None and not isinstance(v, (int, float))
                            or hasattr(v, "year"))
        num_share = _share(vals, lambda v: parse_amount(v) is not None)
        text_share = _share(vals, lambda v: isinstance(v, str) and parse_amount(v) is None)
        for f, syns in SYNONYMS.items():
            s = 0.0
            for syn in syns:
                cs = clean_text(syn)
                if t == cs:
                    s = max(s, 1.0)
                elif len(cs) > 2 and cs in t:
                    s = max(s, 0.7)
                elif len(cs) > 3:
                    try:
                        from rapidfuzz import fuzz
                        if fuzz.ratio(t, cs) >= 88:
                            s = max(s, 0.6)
                    except Exception:
                        pass
            scores[f][j] += s
        # عنوان فيه «مدين/دائن» هو عمود مدين/دائن وليس «مبلغ» عاماً؛ و«نوع ...» وصف لا رقم
        words = set(t.split())
        dc = {clean_text(x) for x in ("مدين", "المدين", "debit", "dr", "دائن", "الدائن", "credit", "cr")}
        if words & dc:
            scores["amount"][j] = 0.0
        if t.startswith("نوع") or t.startswith("المصدر") or t.startswith("مصدر"):
            for f in ("ref", "ref2", "amount", "debit", "credit"):
                scores[f][j] = 0.0
        uniq_all = len({str(v) for v in vals}) if vals else 0
        if vals and len(vals) >= 6 and text_share > 0.7 and uniq_all / len(vals) < 0.3:
            # عمود تصنيف (نوع السند، المستند: «قبض حوالة»…) وليس رقماً مرجعياً
            for f in ("ref", "ref2", "amount", "debit", "credit", "fee"):
                scores[f][j] = 0.0
        elif t.startswith("رقم") and len(t.split()) <= 2 and vals and uniq_all / len(vals) > 0.5:
            scores["ref"][j] = max(scores["ref"][j], 0.65)       # «رقمه»، «رقم القيد»…
        # إشارات من المحتوى
        scores["date"][j] += 0.5 * date_share if date_share > 0.7 else 0
        if num_share > 0.8:
            for f in ("amount", "debit", "credit", "fee"):
                scores[f][j] += 0.15
        if text_share > 0.7:
            scores["party"][j] += 0.1
            scores["desc"][j] += 0.1
        # العملة: قيم قصيرة متكررة
        if vals and text_share > 0.8:
            uniq = {clean_text(v) for v in vals}
            if len(uniq) <= 6 and all(len(u) <= 14 for u in uniq):
                scores["currency"][j] += 0.3
        # المرجع: قيم شبه فريدة
        if vals:
            uniq_ratio = len({str(v) for v in vals}) / len(vals)
            if uniq_ratio > 0.9:
                scores["ref"][j] += 0.15

    mapping: dict[str, Optional[str]] = {f: None for f in FIELDS}
    used: set[int] = set()
    order = ["ref", "date", "amount", "debit", "credit", "currency", "fee", "party", "desc", "ref2"]
    cand = []
    for f in order:
        for j in range(n):
            if scores[f][j] >= 0.6:
                cand.append((scores[f][j], -order.index(f), f, j))
    for s, _, f, j in sorted(cand, reverse=True):
        if mapping[f] is None and j not in used:
            mapping[f] = table.columns[j]
            used.add(j)

    # إن وُجد مدين ودائن فلا حاجة لعمود مبلغ عام إلا إذا كان صريحاً
    if mapping["debit"] and mapping["credit"] and mapping["amount"]:
        a = scores["amount"][table.col_index(mapping["amount"])]
        if a < 1.0:
            used.discard(table.col_index(mapping["amount"]))
            mapping["amount"] = None
    # لو لا مبلغ ولا مدين/دائن: خذ أول عمود رقمي غير مستخدم
    if not mapping["amount"] and not (mapping["debit"] or mapping["credit"]):
        for j in range(n):
            if j in used:
                continue
            vals = _column_values(table, j)
            if vals and _share(vals, lambda v: parse_amount(v) is not None) > 0.9:
                mapping["amount"] = table.columns[j]
                break
    return mapping


# ---------------------------------------------------------------------------
# بناء السجلات
# ---------------------------------------------------------------------------
@dataclass
class SkippedRow:
    side: str
    row: int
    reason: str
    raw: dict


def mapping_problems(mapping: dict[str, Optional[str]]) -> list[str]:
    probs = []
    if not any(mapping.get(f) for f in REQUIRED_ONE_OF_AMOUNT):
        probs.append("لم يتم تحديد عمود المبلغ (أو مدين/دائن).")
    if not any(mapping.get(f) for f in ("ref", "ref2", "desc", "date")):
        probs.append("لم يُحدَّد عمود مرجع أو بيان أو تاريخ؛ ستعتمد المطابقة على المبلغ وبقية الأعمدة فقط.")
    return probs


def build_records(table: RawTable, mapping: dict[str, Optional[str]], side: str,
                  default_currency: str = "") -> tuple[list[Record], list[SkippedRow]]:
    idx = {f: (table.col_index(c) if c and c in table.columns else None) for f, c in mapping.items()}
    roles = {f: c for f, c in mapping.items() if c and c in table.columns}

    def get(row: list[Any], f: str):
        j = idx.get(f)
        return row[j] if j is not None else None

    records: list[Record] = []
    skipped: list[SkippedRow] = []
    for xl_row, row in table.rows:
        raw = {table.columns[j]: row[j] for j in range(len(table.columns))}
        text_blob = " ".join(clean_text(v) for v in row if isinstance(v, str))
        amt = None
        if idx.get("amount") is not None:
            amt = parse_amount(get(row, "amount"))
        if idx.get("debit") is not None or idx.get("credit") is not None:
            d = parse_amount(get(row, "debit")) or Decimal(0)
            c = parse_amount(get(row, "credit")) or Decimal(0)
            if amt is None or (d != 0 or c != 0):
                if amt is None or idx.get("amount") is None:
                    amt = d - c
        ref_cell = get(row, "ref")
        is_total = any(clean_text(w) in text_blob.split() or text_blob.startswith(clean_text(w))
                       for w in TOTAL_WORDS) and not norm_ref(ref_cell)
        if is_total:
            skipped.append(SkippedRow(side, xl_row, "صف إجماليات/ملخص", raw))
            continue
        if amt is None:
            skipped.append(SkippedRow(side, xl_row, "لا يوجد مبلغ صالح", raw))
            continue
        fee = parse_amount(get(row, "fee")) or Decimal(0)
        cur = norm_currency(get(row, "currency"), default_currency)
        ref_raw = ref_cell
        if isinstance(ref_raw, float) and ref_raw.is_integer():
            ref_raw = int(ref_raw)
        records.append(Record(
            side=side, row=xl_row,
            ref="" if ref_raw is None else str(ref_raw).strip(),
            ref_norm=norm_ref(ref_cell),
            ref2="" if get(row, "ref2") is None else str(get(row, "ref2")).strip(),
            date=parse_date(get(row, "date")),
            amount=amt, currency=cur, fee=abs(fee),
            party="" if get(row, "party") is None else str(get(row, "party")).strip(),
            desc="" if get(row, "desc") is None else str(get(row, "desc")).strip(),
            raw=raw, roles=roles,
        ))
    return records, skipped
