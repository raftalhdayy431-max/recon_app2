"""محرك المطابقة والتسوية.

المراحل:
  1) مطابقة بالمرجع (مع معالجة المراجع المتكررة والعمليات المكررة)
  2) مطابقة بالمرجع المذكور داخل البيان / رقم السند
  3) مطابقة ذكية (مبلغ + تاريخ + اسم + تقارب المرجع) للعمليات المتبقية
  4) اكتشاف التكرار
  5) تصنيف المفقودات مع تلميحات تفسيرية

الفرق دائماً = (مبلغ المحاسب) - (مبلغ الصرافة بعد مواءمة الإشارة)
ومجموع كل الفروقات = إجمالي المحاسب - إجمالي الصرافة (تحقق حسابي مدمج).
"""
from __future__ import annotations

import re
from bisect import bisect_left, bisect_right
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Callable, Optional

from .keys import MODE_EXACT, KeyRule, default_rules, extract_tokens
from .loader import Record
from .normalize import clean_text, name_similarity, q2

# ---------------------------------------------------------------------------
STATUS_MATCHED = "متطابقة"
STATUS_DIFF = "مختلفة"
STATUS_MISSING_B = "مفقودة في الصرافة"
STATUS_MISSING_A = "مفقودة في ملف المحاسب"
STATUS_DUP = "مكررة"
STATUS_POSSIBLE = "محتملة المطابقة"

ALL_STATUSES = [STATUS_MATCHED, STATUS_DIFF, STATUS_POSSIBLE, STATUS_MISSING_B,
                STATUS_MISSING_A, STATUS_DUP]
STATUS_ORDER = {s: i for i, s in enumerate(
    [STATUS_DIFF, STATUS_MISSING_B, STATUS_MISSING_A, STATUS_DUP, STATUS_POSSIBLE, STATUS_MATCHED])}

SRC_A = "المحاسب"
SRC_B = "الصرافة"
SRC_BOTH = "الاثنان"

CAT_ROUND = "فروقات تقريب"
CAT_AMOUNT = "فروقات مبالغ (عمليات مرتبطة)"
CAT_FEE = "فروقات مفسّرة بالرسوم"
CAT_SIGN = "إشارة معكوسة (مدين/دائن)"
CAT_CUR = "اختلاف عملة"
CAT_POSSIBLE = "فروقات في المطابقات المحتملة"
CAT_MISS_B = "عمليات عند المحاسب وغير موجودة عند الصرافة"
CAT_MISS_A = "عمليات عند الصرافة وغير موجودة عند المحاسب"
CAT_DUP_A = "مكررة في ملف المحاسب"
CAT_DUP_B = "مكررة في ملف الصرافة"

BRIDGE_ORDER = [CAT_MISS_B, CAT_MISS_A, CAT_DUP_A, CAT_DUP_B, CAT_AMOUNT, CAT_FEE, CAT_SIGN,
                CAT_CUR, CAT_POSSIBLE, CAT_ROUND]


@dataclass
class Settings:
    amount_tol: Decimal = Decimal("0.01")     # تسامح المبلغ (عملة)
    date_tol_days: int = 3                   # فرق تاريخ مقبول للمرجع المتطابق
    fuzzy_days: int = 5                      # نافذة التاريخ للمطابقة الذكية
    possible_pct: float = 0.02               # أعلى نسبة فرق مقبولة للمطابقة المحتملة
    name_threshold: float = 0.80             # حد تشابه الأسماء
    sign_mode: str = "auto"                  # auto / same / opposite / ignore
    strict_currency: bool = True
    use_fuzzy: bool = True
    use_desc_ref: bool = True
    check_party: bool = True
    auto_accept: float = 0.92                # ثقة المطابقة الذكية التي تُعدّ "متطابقة"
    min_possible: float = 0.60               # أدنى ثقة تظهر كمحتملة
    detect_duplicates: bool = True
    key_rules: Optional[list] = None          # قواعد مفاتيح مخصصة (قائمة قواميس)؛ None = تلقائي حسب أعمدة الملفين
    token_min_len: int = 5                    # أقل طول لرقم/رمز يُستخرج من نص العمود (البيان)
    token_max_repeat: int = 3                 # رمز يتكرر في أكثر من هذا العدد من الصفوف لا يُعدّ مفتاحاً
    report_all_columns: bool = False          # تضمين كل أعمدة الملفين الأصلية في التقرير

    def to_dict(self) -> dict:
        d = dict(self.__dict__)
        d["amount_tol"] = str(self.amount_tol)
        return d

    @staticmethod
    def from_dict(d: dict) -> "Settings":
        s = Settings()
        for k, v in d.items():
            if hasattr(s, k):
                if k == "amount_tol":
                    v = Decimal(str(v))
                setattr(s, k, v)
        return s


@dataclass
class MatchRow:
    id: int = 0
    status: str = ""
    method: str = ""
    confidence: int = 0
    a: Optional[Record] = None
    b: Optional[Record] = None
    cur_a: str = ""
    cur_b: str = ""
    amt_a: Optional[Decimal] = None   # بعد المواءمة
    amt_b: Optional[Decimal] = None   # بعد المواءمة
    diff: Decimal = Decimal(0)
    reasons: list[str] = field(default_factory=list)
    action: str = ""
    priority: str = ""
    flags: set = field(default_factory=set)
    explained: bool = False
    source: str = ""
    contribs: list[tuple[str, str, Decimal]] = field(default_factory=list)  # (عملة, فئة, مبلغ)
    note: str = ""   # ملاحظة المراجع (تُحفظ لاحقاً)
    decision: str = ""  # قرار المراجع: "" / معتمد / مرفوض
    members_a: list = field(default_factory=list)   # عمليات ملف المحاسب المجمّعة (مطابقة 1:N)
    members_b: list = field(default_factory=list)   # عمليات ملف الصرافة المجمّعة
    orig: Optional[dict] = None   # لقطة ما قبل اعتماد المراجع (للتراجع عند مسح القرار)

    @property
    def currency(self) -> str:
        return self.cur_a or self.cur_b

    @property
    def ref(self) -> str:
        return (self.a.ref if self.a and self.a.ref else (self.b.ref if self.b else ""))

    @property
    def label(self) -> str:
        """معرّف العملية للعرض: المرجع إن وُجد، وإلا البيان، وإلا رقم الصف."""
        if self.ref:
            return self.ref
        for rec in (self.a, self.b):
            if rec and rec.desc:
                d = " ".join(rec.desc.split())
                return d if len(d) <= 38 else d[:36] + "…"
        rec = self.a or self.b
        return f"صف {rec.row}" if rec else ""

    @property
    def reason_text(self) -> str:
        return " · ".join(self.reasons)


@dataclass
class ReconResult:
    rows: list[MatchRow]
    settings: Settings
    rel: int
    rel_note: str
    count_a: int
    count_b: int
    total_a: dict[str, Decimal]
    total_b: dict[str, Decimal]
    bridge: dict[str, dict[str, Decimal]]     # عملة -> فئة -> مبلغ
    integrity: dict[str, Decimal]             # عملة -> الفرق بين الجسر والفرق الفعلي (يجب = 0)
    counts: dict[str, int]
    warnings: list[str]
    range_a: tuple[Optional[date], Optional[date]]
    range_b: tuple[Optional[date], Optional[date]]
    key_rules: list = field(default_factory=list)   # قواعد المفاتيح المستخدمة فعلاً بالترتيب

    @property
    def currencies(self) -> list[str]:
        cs = set(self.total_a) | set(self.total_b)
        return sorted(cs)

    @property
    def match_rate(self) -> float:
        n = max(self.count_a, self.count_b, 1)
        return self.counts.get(STATUS_MATCHED, 0) / n

    def net_diff(self, cur: str) -> Decimal:
        return self.total_a.get(cur, Decimal(0)) - self.total_b.get(cur, Decimal(0))


# ---------------------------------------------------------------------------
def _eff(rec: Record, rel: int, mode: str) -> Decimal:
    if mode == "ignore":
        return abs(rec.amount)
    return rec.amount if rec.side == "A" else rec.amount * rel


def _days(a: Optional[date], b: Optional[date]) -> Optional[int]:
    if a is None or b is None:
        return None
    return abs((a - b).days)


def _fmt(d: Decimal) -> str:
    return f"{q2(d):,.2f}"


def _digits_key(d: Decimal) -> str:
    return str(q2(abs(d))).replace(".", "")


def _explain_amount(mag: Decimal, a: Record, b: Record, S: Settings) -> tuple[list[str], bool, set]:
    """تفسير فرق المبلغ. mag = |A| - |B|."""
    reasons: list[str] = []
    flags: set = set()
    explained = False
    am = abs(mag)
    if mag > 0:
        reasons.append(f"مبلغ المحاسب أكبر من مبلغ الصرافة بمقدار {_fmt(am)}")
    else:
        reasons.append(f"مبلغ الصرافة أكبر من مبلغ المحاسب بمقدار {_fmt(am)}")

    for who, fee in (("الصرافة", b.fee), ("المحاسب", a.fee)):
        if fee and fee > 0 and abs(am - fee) <= S.amount_tol:
            reasons.append(f"الفرق يساوي الرسوم/العمولة المسجلة في ملف {who} ({_fmt(fee)})")
            explained = True
            flags.add("fee")
            break

    if not explained:
        lo, hi = sorted([abs(a.amount), abs(b.amount)])
        if lo > 0:
            ratio = hi / lo
            for r in (10, 100, 1000):
                if abs(ratio - r) / r < Decimal("0.0005"):
                    reasons.append("احتمال خطأ في موضع الفاصلة العشرية (النسبة بين المبلغين "
                                   f"{r}×)")
                    flags.add("decimal")
                    break
            sa, sb = _digits_key(a.amount), _digits_key(b.amount)
            if "decimal" not in flags and sa != sb and len(sa) >= 2 and sorted(sa) == sorted(sb):
                reasons.append("احتمال تبديل أرقام عند الإدخال (نفس الأرقام بترتيب مختلف)")
                flags.add("transpose")
            if "decimal" not in flags and "transpose" not in flags:
                p = am / hi
                if am < 1:
                    reasons.append("فرق تقريب صغير")
                    flags.add("rounding")
                elif p <= Decimal("0.03"):
                    reasons.append(f"الفرق يمثل {float(p) * 100:.2f}% من المبلغ — قد يكون عمولة أو "
                                   "سعر صرف غير موثّق")
                    flags.add("pct")
                else:
                    reasons.append("فرق غير مفسّر بالبيانات المتاحة")
                    flags.add("unexplained")
    return reasons, explained, flags


class Engine:
    def __init__(self, settings: Optional[Settings] = None,
                 progress: Optional[Callable[[int, str], None]] = None):
        self.S = settings or Settings()
        self.progress = progress or (lambda p, m: None)
        self.rel = 1

    # ------------------------------------------------------------------ عام
    def eff(self, rec: Record) -> Decimal:
        return _eff(rec, self.rel, self.S.sign_mode)

    def _row_from_pair(self, a: Record, b: Record, method: str, conf: int) -> MatchRow:
        S = self.S
        if bool(a.currency) != bool(b.currency):      # عملة غائبة في أحد الطرفين: ترث عملة نظيرها
            if a.currency:
                b.currency = a.currency
            else:
                a.currency = b.currency
        ea, eb = self.eff(a), self.eff(b)
        row = MatchRow(a=a, b=b, method=method, confidence=conf,
                       cur_a=a.currency, cur_b=b.currency, amt_a=ea, amt_b=eb,
                       diff=ea - eb, source=SRC_BOTH)
        flags: set = set()
        reasons: list[str] = []

        sign_flag = (S.sign_mode != "ignore" and ea != 0 and eb != 0 and (ea > 0) != (eb > 0))
        if sign_flag:
            flags.add("sign")
            reasons.append("إشارة العملية معكوسة بين الملفين (مدين/دائن)")

        mag = abs(a.amount) - abs(b.amount)
        explained = False
        if abs(mag) > S.amount_tol:
            flags.add("amount")
            r, explained, f2 = _explain_amount(mag, a, b, S)
            reasons.extend(r)
            flags |= f2

        if S.strict_currency and a.currency and b.currency and a.currency != b.currency:
            flags.add("currency")
            reasons.append(f"اختلاف العملة: المحاسب {a.currency} مقابل الصرافة {b.currency}")

        d = _days(a.date, b.date)
        if d is not None and d > S.date_tol_days:
            flags.add("date")
            reasons.append(f"فرق التاريخ {d} يوم (المحاسب {a.date:%Y-%m-%d} / الصرافة {b.date:%Y-%m-%d})")

        if S.check_party and a.party and b.party:
            sim = name_similarity(a.party, b.party)
            if sim < 0.35:
                flags.add("party")
                reasons.append("تنبيه: اسم الطرف مختلف بين الملفين")

        hard = flags & {"amount", "sign", "currency", "date"}
        row.flags = flags
        row.explained = explained and (flags & {"sign", "currency", "date"}) == set()
        if hard:
            row.status = STATUS_DIFF
        else:
            row.status = STATUS_MATCHED
            if not reasons:
                reasons.append("تطابق تام" if method not in ("مطابقة ذكية", "ربط يدوي") else "تطابق تام في البيانات المتاحة")
        row.reasons = reasons

        # مساهمة الجسر الحسابي
        cur = a.currency or b.currency
        if "currency" in flags:
            row.contribs = [(a.currency, CAT_CUR, ea), (b.currency, CAT_CUR, -eb)]
        else:
            if "sign" in flags:
                cat = CAT_SIGN
            elif "amount" in flags and row.explained and "fee" in flags:
                cat = CAT_FEE
            elif "amount" in flags:
                cat = CAT_AMOUNT
            else:
                cat = CAT_ROUND
            if method == "مطابقة ذكية" and row.status != STATUS_MATCHED:
                cat = CAT_POSSIBLE
            row.contribs = [(cur, cat, row.diff)]
        return row

    # ------------------------------------------------------------------ التشغيل
    def run(self, recs_a: list[Record], recs_b: list[Record]) -> ReconResult:
        S = self.S
        warnings: list[str] = []
        self.progress(5, "تجهيز السجلات")

        unmatched_a: dict[str, Record] = {r.uid: r for r in recs_a}
        unmatched_b: dict[str, Record] = {r.uid: r for r in recs_b}
        self._unify_currency(recs_a, recs_b, warnings)
        self._prepare_words(recs_a, recs_b)
        self._check_ref_comparable(recs_a, recs_b, warnings)
        rows: list[MatchRow] = []
        raw_pairs: list[tuple[Record, Record, str, int]] = []
        dup_rows: list[MatchRow] = []

        # ---- المرحلة 1: قواعد المفاتيح بالأولوية (البيان ثم المرجع ثم بقية الأعمدة ...) ----
        dup_pending: list[tuple[Record, str]] = []
        rules = self._resolve_rules(recs_a, recs_b, warnings)
        self.rules_used = rules
        if not rules:
            warnings.append("لا توجد أعمدة مفاتيح مشتركة بين الملفين؛ اعتُمدت المطابقة الذكية (مبلغ/تاريخ/اسم) فقط.")
        for k, rule in enumerate(rules):
            self.progress(15 + int(30 * k / max(len(rules), 1)), f"المطابقة بالعمود: {rule.name}")
            self._key_stage(rule, recs_a, recs_b, unmatched_a, unmatched_b, raw_pairs, dup_pending)

        # ---- تحديد علاقة الإشارة ----
        self._decide_sign([(a, b) for a, b, *_ in raw_pairs], warnings)

        for a, b, m, c, note in raw_pairs:
            row = self._row_from_pair(a, b, m, c)
            if row.status == STATUS_DIFF and c == 100:
                row.confidence = 95
            if note:
                row.reasons.insert(0, note)
            rows.append(row)
        for rec, why in dup_pending:
            dup_rows.append(self._dup_row(rec, why))

        # ---- المرحلة 3: المطابقة الذكية ----
        if S.use_fuzzy and unmatched_a and unmatched_b:
            self.progress(55, "المطابقة الذكية (مبلغ + تاريخ + اسم)")
            self._fuzzy_stage(unmatched_a, unmatched_b, rows)

        # ---- المرحلة 3ب: مطابقة 1:N (قيد واحد = مجموع عدة عمليات، كالرسوم الشهرية) ----
        if unmatched_a and unmatched_b:
            self.progress(70, "البحث عن قيود تساوي مجموع عمليات")
            self._group_stage(unmatched_a, unmatched_b, rows)

        # ---- المرحلة 4: تكرار بلا مرجع + المفقودات ----
        self.progress(80, "تصنيف المفقودات والمكررات")
        if S.detect_duplicates:
            self._dups_unmatched(unmatched_a, rows, dup_rows, "A")
            self._dups_unmatched(unmatched_b, rows, dup_rows, "B")
        rows.extend(dup_rows)

        range_a = self._range(recs_a)
        range_b = self._range(recs_b)
        self._missing_rows(unmatched_a, unmatched_b, rows, range_a, range_b)

        self._mark_storno(rows)
        # ---- إجراءات وأولويات ----
        for r in rows:
            self._finalize(r)
        rows.sort(key=lambda r: (STATUS_ORDER[r.status],
                                 -(abs(r.diff)),
                                 (r.a.row if r.a else 10**9), (r.b.row if r.b else 10**9)))
        for i, r in enumerate(rows, 1):
            r.id = i

        self.progress(92, "حساب الفروقات والتحقق الحسابي")
        res = self._summarize(rows, recs_a, recs_b, warnings, range_a, range_b)
        res.key_rules = list(rules)
        self.progress(100, "اكتملت المطابقة")
        return res

    # ------------------------------------------------------------------ مساعدات
    @staticmethod
    def _range(recs: list[Record]):
        ds = [r.date for r in recs if r.date]
        return (min(ds), max(ds)) if ds else (None, None)

    def _decide_sign(self, pairs, warnings):
        S = self.S
        if S.sign_mode == "same":
            self.rel, self.rel_note = 1, "اتجاه الإشارة: نفس الاتجاه في الملفين (إعداد يدوي)"
        elif S.sign_mode == "opposite":
            self.rel, self.rel_note = -1, "اتجاه الإشارة: معكوس بين الملفين (إعداد يدوي)"
        elif S.sign_mode == "ignore":
            self.rel, self.rel_note = 1, "تجاهل الإشارات: المقارنة بالقيمة المطلقة"
        else:
            same = opp = 0
            for a, b, *_ in pairs:
                if a.amount == 0 or b.amount == 0:
                    continue
                if (a.amount > 0) == (b.amount > 0):
                    same += 1
                else:
                    opp += 1
            if opp > same:
                self.rel = -1
                self.rel_note = (f"اكتُشف تلقائياً أن إشارات الصرافة معكوسة عن المحاسب "
                                 f"({opp} عملية معكوسة مقابل {same})")
            else:
                self.rel = 1
                self.rel_note = (f"اكتُشف تلقائياً أن الإشارات بنفس الاتجاه "
                                 f"({same} مقابل {opp} معكوسة)" if (same or opp)
                                 else "لا توجد عمليات مرجعية كافية لتحديد الإشارة؛ افتُرض نفس الاتجاه")

    def _is_dup_of(self, r: Record, others: list[Record]) -> bool:
        for o in others:
            if (abs(o.amount - r.amount) <= self.S.amount_tol and o.currency == r.currency
                    and (o.date == r.date or o.date is None or r.date is None)):
                return True
        return False

    def _dup_row(self, r: Record, why: str) -> MatchRow:
        e = self.eff(r)
        row = MatchRow(status=STATUS_DUP, method="كشف التكرار", confidence=90,
                       cur_a=r.currency if r.side == "A" else "",
                       cur_b=r.currency if r.side == "B" else "",
                       reasons=[why])
        if r.side == "A":
            row.a, row.amt_a, row.diff, row.source = r, e, e, SRC_A
            row.contribs = [(r.currency, CAT_DUP_A, e)]
        else:
            row.b, row.amt_b, row.diff, row.source = r, e, -e, SRC_B
            row.contribs = [(r.currency, CAT_DUP_B, -e)]
        return row

    def _dups_unmatched(self, un: dict[str, Record], rows, dup_rows, side):
        """تكرار بين السجلات غير المطابقة (أو مماثلة لسجل مطابَق) بلا مرجع."""
        matched_recs = [getattr(r, "a" if side == "A" else "b") for r in rows
                        if getattr(r, "a" if side == "A" else "b") is not None]
        def key(r: Record):
            return (r.ref_norm, q2(abs(r.amount)), r.currency, r.date,
                    re.sub(r"\s+", " ", r.party.strip().lower()), re.sub(r"\s+", " ", (r.desc or "").strip().lower()))
        seen: dict = {}
        for m in matched_recs:
            if m.ref_norm or m.party or m.desc:
                seen.setdefault(key(m), m)
        for uid, r in sorted(un.items(), key=lambda kv: kv[1].row):
            if not (r.ref_norm or r.party or r.desc):
                continue
            k = key(r)
            if k in seen:
                side_name = "المحاسب" if side == "A" else "الصرافة"
                dup_rows.append(self._dup_row(
                    r, f"تطابق كامل (المرجع/المبلغ/التاريخ/الاسم) مع العملية في الصف {seen[k].row} من ملف {side_name}"))
                un[uid] = None  # علامة للحذف
            else:
                seen[k] = r
        for uid in [u for u, v in un.items() if v is None]:
            del un[uid]

    # ---- قواعد المفاتيح الديناميكية ----
    def _resolve_rules(self, recs_a, recs_b, warnings) -> list[KeyRule]:
        virt = {"ref": "@ref", "ref2": "@ref2", "desc": "@desc"}

        def side_info(recs):
            if not recs:
                return [], {}
            if recs[0].roles or recs[0].raw:
                return list(recs[0].raw.keys()), dict(recs[0].roles)
            return list(virt.values()), dict(virt)      # سجلات مبنية برمجياً بلا أعمدة أصلية

        cols_a, roles_a = side_info(recs_a)
        cols_b, roles_b = side_info(recs_b)
        if self.S.key_rules:
            out = []
            for d in self.S.key_rules:
                r = KeyRule.from_dict(d)
                if not r.enabled:
                    continue
                a = [c for c in r.a if c in cols_a]
                b = [c for c in r.b if c in cols_b]
                for side, want, have in (("المحاسب", r.a, a), ("الصرافة", r.b, b)):
                    miss = [c for c in want if c not in have]
                    if miss and (a or b):
                        warnings.append(f"القاعدة «{r.name}»: العمود {'، '.join(miss)} غير موجود في ملف {side} فتم تجاهله.")
                if a and b:
                    out.append(KeyRule(r.name, a, b, r.mode))
            return out
        auto = default_rules(cols_a, roles_a, cols_b, roles_b)
        return auto if self.S.use_desc_ref else [r for r in auto if r.mode == MODE_EXACT]

    def _unify_currency(self, recs_a, recs_b, warnings):
        """ملف بلا عمود عملة (أو فارغ) يرث عملة الملف الآخر؛ حتى لا ينفصل الإجمالي بين «بلا عملة» و«YER»."""
        ca = Counter(r.currency for r in recs_a if r.currency)
        cb = Counter(r.currency for r in recs_b if r.currency)
        for mine, other_c, name_mine, name_other in ((recs_a, cb, "المحاسب", "الصرافة"),
                                                     (recs_b, ca, "الصرافة", "المحاسب")):
            have = Counter(r.currency for r in mine if r.currency)
            if other_c and (not have or sum(have.values()) < 0.2 * len(mine)):
                cur = other_c.most_common(1)[0][0]
                for r in mine:
                    if not r.currency:
                        r.currency = cur
                note = f"ملف {name_mine} بلا عملة؛ اعتُمدت عملة ملف {name_other} ({cur})"
                if len(other_c) > 1:
                    note += " — الملف الآخر متعدد العملات فراجع العملات يدوياً"
                warnings.append(note)

    def _check_ref_comparable(self, recs_a, recs_b, warnings):
        """أرقام مرجعية من نظامين مختلفين (سند المحاسب ↔ سند الصرافة) لا تُقارَن ببعضها."""
        ra = {r.ref_norm for r in recs_a if r.ref_norm}
        rb = {r.ref_norm for r in recs_b if r.ref_norm}
        self._ref_ok = True
        if ra and rb:
            common = len(ra & rb)
            if common < max(2, 0.02 * min(len(ra), len(rb))):
                self._ref_ok = False
                warnings.append("عمودا «المرجع» في الملفين من نظامين مختلفين (لا قيم مشتركة)؛ "
                                "لم يُعتمد المرجع في المطابقة الذكية، والاعتماد على البيان والمبلغ والتاريخ.")

    _CUE = r"(?:س\s*\.?\s*ص|س\s*\.?\s*ق|سند(?:\s*(?:صرف|قبض))?(?:\s*رقم)?|قيد|voucher|vch)\s*[:\-]?\s*"

    def _cue_hit(self, tokens, a: Record, b: Record, rule) -> bool:
        """الرقم القصير مسبوق بكلمة سند/«س.ص»: إحالة صريحة إلى رقم سند الطرف الآخر."""
        for rec, cols in ((a, rule.a), (b, rule.b)):
            for c in cols:
                val = getattr(rec, c[1:], None) if c.startswith("@") else rec.raw.get(c)
                if not isinstance(val, str):
                    continue
                txt = clean_text(val)
                for t in tokens:
                    if re.search(self._CUE + re.escape(t) + r"(?!\d)", txt, re.I):
                        return True
        return False

    # ---- كلمات البيان المميِّزة (أسماء/تفاصيل) تُستخدم كدليل تأكيد عند غياب عمود الاسم ----
    _WORD = re.compile(r"[^\W\d_]{3,}", re.UNICODE)

    def _prepare_words(self, recs_a, recs_b):
        self._dw: dict[str, frozenset] = {}
        for recs in (recs_a, recs_b):
            words = {}
            for r in recs:
                txt = " ".join(x for x in (r.desc, r.party) if x)
                words[r.uid] = set(self._WORD.findall(clean_text(txt))) if txt else set()
            n = len(recs)
            df = Counter(w for ws in words.values() for w in ws)
            lim = max(2, int(n * 0.25))
            for uid, ws in words.items():
                self._dw[uid] = frozenset(w for w in ws if df[w] <= lim)

    def _word_overlap(self, a: Record, b: Record) -> Optional[float]:
        """نسبة الكلمات المميِّزة المشتركة (None = لا دليل نصي كافٍ لأحد الطرفين)."""
        wa, wb = self._dw.get(a.uid), self._dw.get(b.uid)
        if not wa or not wb:
            return None
        small, big = (wa, wb) if len(wa) <= len(wb) else (wb, wa)
        hit = len(small & big)
        rest_s, rest_b = list(small - big), list(big - small)
        if rest_s and rest_b:     # تقارب إملائي (عبدالغني/الغني، حسين/حسن…)
            from rapidfuzz import fuzz
            for w in rest_s:
                best = max(rest_b, key=lambda x: fuzz.ratio(w, x))
                if fuzz.ratio(w, best) >= 82:
                    hit += 1
                    rest_b.remove(best)
        return hit / len(small)

    def _key_stage(self, rule: KeyRule, recs_a, recs_b, ua, ub, pairs, dup_pending):
        S = self.S
        minlen = 1 if rule.mode == MODE_EXACT else min(3, S.token_min_len)

        def toks(rec, cols):
            d: dict[str, str] = {}
            for c in cols:
                val = getattr(rec, c[1:], None) if c.startswith("@") else rec.raw.get(c)
                for t in extract_tokens(val, rule.mode, minlen):
                    d.setdefault(t, c)
            return d

        tok_a = {r.uid: toks(r, rule.a) for r in recs_a}
        tok_b = {r.uid: toks(r, rule.b) for r in recs_b}
        if rule.mode != MODE_EXACT:   # رموز كثيرة التكرار (كود بنك، كلمة شائعة) لا تصلح مفتاحاً
            for tk in (tok_a, tok_b):
                freq = Counter(t for d in tk.values() for t in d)
                for d in tk.values():
                    for t in [t for t in d if freq[t] > S.token_max_repeat]:
                        del d[t]

        idx_b: dict[str, list[Record]] = defaultdict(list)
        for r in recs_b:
            if r.uid in ub:
                for t in tok_b[r.uid]:
                    idx_b[t].append(r)
        cands: dict[tuple[str, str], list[str]] = defaultdict(list)
        recs_by_uid = {}
        for a in recs_a:
            if a.uid not in ua:
                continue
            for t in tok_a[a.uid]:
                for b in idx_b.get(t, ()):
                    cands[(a.uid, b.uid)].append(t)
                    recs_by_uid[b.uid] = b
            recs_by_uid[a.uid] = a
        if rule.mode != MODE_EXACT and S.token_min_len > 3:
            # أرقام قصيرة (رقم سند 3706): لا تكفي وحدها؛ تُقبل فقط مع تطابق المبلغ ودليل نصي (اسم/تفصيل مشترك)
            for key in list(cands):
                ts = cands[key]
                good = [t for t in ts if len(t) >= S.token_min_len]
                if good:
                    cands[key] = good
                    continue
                a, b = recs_by_uid[key[0]], recs_by_uid[key[1]]
                ov = self._word_overlap(a, b)
                same_amt = abs(abs(a.amount) - abs(b.amount)) <= S.amount_tol
                if same_amt and (ov is None or ov >= 0.34 or self._cue_hit(ts, a, b, rule)):
                    continue
                del cands[key]
        if rule.mode != MODE_EXACT:
            # رمز واحد ضعيف (أرقام قصيرة) مع اختلاف المبلغ ولا دليل نصي → تصادف لا ارتباط
            for key in list(cands):
                ts = cands[key]
                a, b = recs_by_uid[key[0]], recs_by_uid[key[1]]
                if (len(ts) == 1 and len(ts[0]) <= 6 and ts[0].isdigit()
                        and abs(abs(a.amount) - abs(b.amount)) > S.amount_tol
                        and (self._word_overlap(a, b) or 0) < 0.34):
                    del cands[key]
        if not cands:
            return

        scored = []
        for (au, bu), ts in cands.items():
            a, b = recs_by_uid[au], recs_by_uid[bu]
            cm = 1 if (a.currency and b.currency and a.currency != b.currency) else 0
            dd = _days(a.date, b.date)
            scored.append(((-len(ts), cm, abs(abs(a.amount) - abs(b.amount)), dd if dd is not None else 9999),
                           a.row, b.row, au, bu, ts))
        scored.sort(key=lambda x: (x[0], x[1], x[2]))

        ambiguous: set[tuple[str, str]] = set()
        if rule.mode != MODE_EXACT:
            by_a, by_b = defaultdict(list), defaultdict(list)
            for s in scored:
                by_a[s[3]].append(s)
                by_b[s[4]].append(s)
            for grp in (by_a, by_b):
                for lst in grp.values():
                    if len(lst) > 1 and lst[0][0] == lst[1][0]:
                        ambiguous.add((lst[0][3], lst[0][4]))
                        ambiguous.add((lst[1][3], lst[1][4]))

        used_a, used_b = set(), set()
        new_pairs = []
        conf = 100 if rule.mode == MODE_EXACT else 90
        for key, _ra, _rb, au, bu, ts in scored:
            if au in used_a or bu in used_b or (au, bu) in ambiguous:
                continue
            a, b = recs_by_uid[au], recs_by_uid[bu]
            used_a.add(au)
            used_b.add(bu)
            note = ""
            if rule.mode != MODE_EXACT:
                t = ts[0]
                note = (f"رُبطت العمليتان لأن القيمة «{t}» مشتركة بين عمود «{tok_a[au][t]}» (ملف المحاسب) "
                        f"وعمود «{tok_b[bu][t]}» (ملف الصرافة)")
            new_pairs.append((a, b, rule.name, conf, note))
        pairs.extend(new_pairs)
        for a, b, *_ in new_pairs:
            ua.pop(a.uid, None)
            ub.pop(b.uid, None)

        # ---- تكرار: سجل لم يُطابَق وله مثيل مطابَق/مماثل في نفس الملف بنفس المفتاح ----
        if S.detect_duplicates:
            side_a, side_b = defaultdict(list), defaultdict(list)
            for r in recs_a:
                for t in tok_a[r.uid]:
                    side_a[t].append(r)
            for r in recs_b:
                for t in tok_b[r.uid]:
                    side_b[t].append(r)
            in_cands_a = {k[0] for k in cands}
            in_cands_b = {k[1] for k in cands}
            for uid in in_cands_a:
                r = recs_by_uid[uid]
                if uid in ua and uid not in used_a:
                    for t in tok_a[uid]:
                        if self._is_dup_of(r, [x for x in side_a[t] if x.uid != uid]):
                            dup_pending.append((r, "نفس المفتاح والمبلغ والتاريخ مسجّل أكثر من مرة في ملف المحاسب"))
                            ua.pop(uid, None)
                            break
            for uid in in_cands_b:
                r = recs_by_uid[uid]
                if uid in ub and uid not in used_b:
                    for t in tok_b[uid]:
                        if self._is_dup_of(r, [x for x in side_b[t] if x.uid != uid]):
                            dup_pending.append((r, "نفس المفتاح والمبلغ والتاريخ مسجّل أكثر من مرة في ملف الصرافة"))
                            ub.pop(uid, None)
                            break

    # ---- المطابقة الذكية ----
    def _score(self, a: Record, b: Record) -> Optional[float]:
        S = self.S
        if S.strict_currency and a.currency and b.currency and a.currency != b.currency:
            return None
        ma = abs(a.amount) - abs(b.amount)
        am = abs(ma)
        hi = max(abs(a.amount), abs(b.amount))
        if am <= S.amount_tol:
            amt_s = 1.0
        else:
            fee_ok = any(f and f > 0 and abs(am - f) <= S.amount_tol for f in (a.fee, b.fee))
            if fee_ok:
                amt_s = 0.85
            else:
                p = float(am / hi) if hi else 1.0
                if p > S.possible_pct:
                    return None
                amt_s = max(0.3, 0.7 * (1 - p / max(S.possible_pct, 1e-9)))

        # المرجع
        ref_sim = None
        if getattr(self, "_ref_ok", True) and a.ref_norm and b.ref_norm and len(a.ref_norm) >= 4 and len(b.ref_norm) >= 4:
            try:
                from rapidfuzz import fuzz
                ref_sim = fuzz.ratio(a.ref_norm, b.ref_norm) / 100.0
            except Exception:
                ref_sim = 1.0 if a.ref_norm == b.ref_norm else 0.0

        d = _days(a.date, b.date)
        if d is None:
            date_s = 0.5
        elif d > S.fuzzy_days:
            if ref_sim is not None and ref_sim >= 0.85 and amt_s >= 0.85:
                date_s = 0.2
            else:
                return None
        else:
            date_s = 1 - d / (S.fuzzy_days + 1)

        if a.party and b.party:
            sim = name_similarity(a.party, b.party)
            name_s = sim
        else:
            ov = self._word_overlap(a, b)
            name_s = 0.5 if ov is None else (0.2 + 0.8 * ov if ov > 0 else 0.0)

        ref_s = 0.5 if ref_sim is None else ref_sim
        if ref_sim is None:
            # لا مرجع عند أحد الطرفين: لا يُعاقَب التطابق بدرجة «محايدة» للمرجع
            conf = (0.40 * amt_s + 0.25 * date_s + 0.25 * name_s) / 0.90
        else:
            conf = 0.40 * amt_s + 0.25 * date_s + 0.25 * name_s + 0.10 * ref_s
        if ref_sim is not None:
            if ref_sim < 0.6:
                conf *= 0.6
            elif ref_sim >= 0.85:
                conf = min(1.0, conf + 0.10)
        return conf

    def _fuzzy_stage(self, ua, ub, rows):
        S = self.S
        b_list = sorted(ub.values(), key=lambda r: abs(r.amount))
        b_keys = [abs(r.amount) for r in b_list]
        max_fee_b = max((r.fee for r in b_list), default=Decimal(0))

        edges = []
        for a in ua.values():
            x = abs(a.amount)
            pad = x * Decimal(str(S.possible_pct)) + a.fee + max_fee_b + S.amount_tol
            lo = bisect_left(b_keys, x - pad)
            hi = bisect_right(b_keys, x + pad)
            for b in b_list[lo:hi]:
                c = self._score(a, b)
                if c is not None and c >= S.min_possible:
                    edges.append((c, a.uid, b.uid))
        edges.sort(key=lambda e: -e[0])

        by_a, by_b = defaultdict(list), defaultdict(list)
        for c, au, bu in edges:
            by_a[au].append(c)
            by_b[bu].append(c)

        used_a, used_b = set(), set()
        for c, au, bu in edges:
            if au in used_a or bu in used_b:
                continue
            a, b = ua[au], ub[bu]
            # غموض: مرشح آخر قريب في الثقة
            alt_a = sorted(by_a[au], reverse=True)
            alt_b = sorted(by_b[bu], reverse=True)
            ambiguous = (len(alt_a) > 1 and alt_a[1] >= c - 0.03) or (len(alt_b) > 1 and alt_b[1] >= c - 0.03)
            conf = min(c, 0.98) * (0.9 if ambiguous else 1.0)
            row = self._row_from_pair(a, b, "مطابقة ذكية", int(round(conf * 100)))
            refs_conflict = bool(getattr(self, "_ref_ok", True) and a.ref_norm and b.ref_norm and a.ref_norm != b.ref_norm)
            if row.status == STATUS_MATCHED and conf >= S.auto_accept and not refs_conflict and not ambiguous:
                row.reasons = ["مطابقة ذكية: تطابق المبلغ والتاريخ والاسم بدون مرجع مشترك"]
            else:
                row.status = STATUS_POSSIBLE
                row.contribs = [(cc, CAT_POSSIBLE if cat in (CAT_ROUND, CAT_AMOUNT, CAT_FEE, CAT_SIGN) else cat, m)
                                for cc, cat, m in row.contribs]
                notes = []
                if refs_conflict:
                    from rapidfuzz import fuzz
                    sim = fuzz.ratio(a.ref_norm, b.ref_norm)
                    if sim >= 85:
                        notes.append(f"المرجعان متقاربان جداً ({a.ref} ↔ {b.ref}) — احتمال خطأ طباعي")
                    else:
                        notes.append(f"المرجعان مختلفان ({a.ref} ↔ {b.ref}) رغم تقارب باقي البيانات")
                elif not a.ref_norm or not b.ref_norm:
                    notes.append("لا يوجد مرجع في أحد الطرفين؛ الاعتماد على المبلغ والتاريخ والاسم")
                if ambiguous:
                    notes.append("يوجد أكثر من مرشح بنفس الدرجة تقريباً — يلزم اختيار يدوي")
                row.reasons = notes + [r for r in row.reasons if not r.startswith("تطابق تام")]
            rows.append(row)
            used_a.add(au)
            used_b.add(bu)
        for u in used_a:
            ua.pop(u, None)
        for u in used_b:
            ub.pop(u, None)

    # ---- المفقودات ----
    def _missing_rows(self, ua, ub, rows, range_a, range_b):
        def hints(r: Record, others: dict[str, Record], other_range, other_name):
            out = []
            for o in others.values():
                if (abs(abs(o.amount) - abs(r.amount)) <= self.S.amount_tol
                        and (not r.currency or not o.currency or r.currency == o.currency)):
                    d = _days(r.date, o.date)
                    extra = f" بفارق {d} يوم" if d is not None else ""
                    out.append(f"يوجد في ملف {other_name} عملية بنفس المبلغ (الصف {o.row}) لكن لم تُربط{extra}")
                    break
            lo, hi = other_range
            if r.date and lo and hi and (r.date < lo or r.date > hi):
                out.append(f"تاريخ العملية خارج الفترة التي يغطيها ملف {other_name} "
                           f"({lo:%Y-%m-%d} → {hi:%Y-%m-%d}) — قد تكون ضمن فترة مختلفة")
            return out

        for r in ua.values():
            e = self.eff(r)
            row = MatchRow(status=STATUS_MISSING_B, method="لا يوجد مقابل", confidence=0, a=r,
                           cur_a=r.currency, amt_a=e, diff=e, source=SRC_A,
                           reasons=["غير موجودة في ملف الصرافة"] + hints(r, ub, range_b, "الصرافة"),
                           contribs=[(r.currency, CAT_MISS_B, e)])
            rows.append(row)
        for r in ub.values():
            e = self.eff(r)
            row = MatchRow(status=STATUS_MISSING_A, method="لا يوجد مقابل", confidence=0, b=r,
                           cur_b=r.currency, amt_b=e, diff=-e, source=SRC_B,
                           reasons=["غير موجودة في ملف المحاسب"] + hints(r, ua, range_a, "المحاسب"),
                           contribs=[(r.currency, CAT_MISS_A, -e)])
            rows.append(row)

    # ---- مطابقة 1:N ----
    @staticmethod
    def _template(r: Record) -> str:
        words = re.findall(r"[^\W\d_]{2,}", clean_text(r.desc or ""))
        return " ".join(words[:2])

    def _group_stage(self, ua, ub, rows):
        S = self.S
        tol = S.amount_tol
        for single_side, pool_side in (("A", "B"), ("B", "A")):
            singles = ua if single_side == "A" else ub
            pool = ub if single_side == "A" else ua
            groups: dict[tuple, list[Record]] = defaultdict(list)
            for r in pool.values():
                if self.eff(r) != 0:
                    groups[(self._template(r), self.eff(r) > 0, r.currency)].append(r)
            for sr in sorted(list(singles.values()), key=lambda r: -abs(r.amount)):
                if sr.uid not in singles:
                    continue
                es = self.eff(sr)
                for (tmpl, pos, cur), members in groups.items():
                    if len(members) < 2 or not tmpl or (es > 0) != pos or (sr.currency and cur and sr.currency != cur):
                        continue
                    live = [m for m in members if m.uid in pool]
                    if len(live) < 2:
                        continue
                    pick = None
                    tot = sum((self.eff(m) for m in live), Decimal(0))
                    if abs(tot - es) <= tol:
                        pick = live
                    elif len(live) <= 16 and abs(es) < abs(tot):
                        vals = [self.eff(m) for m in live]
                        for mask in range(3, 1 << len(live)):
                            if mask & (mask - 1) == 0:
                                continue
                            sm = sum((v for k, v in enumerate(vals) if mask >> k & 1), Decimal(0))
                            if abs(sm - es) <= tol:
                                pick = [m for k, m in enumerate(live) if mask >> k & 1]
                                break
                    if not pick:
                        continue
                    agg = self._aggregate(pick, single=sr)
                    a, b = (sr, agg) if single_side == "A" else (agg, sr)
                    row = self._row_from_pair(a, b, "تجميع (1:N)", 93)
                    for m in pick:                    # العملة الموروثة تنتقل للأعضاء الأصليين
                        m.currency = agg.currency
                    n = len(pick)
                    rows_txt = "، ".join(str(m.row) for m in pick[:12]) + ("…" if n > 12 else "")
                    row.reasons.insert(0, f"قيد واحد (الصف {sr.row}) يساوي مجموع {n} عملية في الملف الآخر "
                                          f"(الصفوف {rows_txt}) بالمبلغ نفسه")
                    if single_side == "A":
                        row.members_b = list(pick)
                    else:
                        row.members_a = list(pick)
                    rows.append(row)
                    singles.pop(sr.uid, None)
                    for m in pick:
                        pool.pop(m.uid, None)
                    break

    def _aggregate(self, members: list[Record], single: Record) -> Record:
        first = members[0]
        descs = "؛ ".join(" ".join((m.desc or "").split())[:40] for m in members[:3])
        return Record(side=first.side, row=first.row, ref="", ref_norm="", ref2="", date=single.date,
                      amount=sum((m.amount for m in members), Decimal(0)), currency=first.currency,
                      fee=sum((m.fee for m in members), Decimal(0)), party="",
                      desc=f"مجموع {len(members)} عملية: {descs}…", raw=dict(first.raw), roles=dict(first.roles),
                      uid=f"{first.side}G{first.row}")

    # ---- قيد وعكسه في نفس الملف ----
    def _mark_storno(self, rows):
        for st in (STATUS_MISSING_A, STATUS_MISSING_B):
            side = [r for r in rows if r.status == st]
            by_amt: dict[tuple, list[MatchRow]] = defaultdict(list)
            for r in side:
                rec = r.a or r.b
                by_amt[(q2(abs(rec.amount)), rec.currency)].append(r)
            for lst in by_amt.values():
                pos = [r for r in lst if (r.a or r.b).amount > 0]
                neg = [r for r in lst if (r.a or r.b).amount < 0]
                for p, n in zip(pos, neg):
                    for x, y in ((p, n), (n, p)):
                        rec_y = y.a or y.b
                        x.flags.add("storno")
                        x.reasons.insert(0, f"له عملية معاكسة بنفس المبلغ في الملف نفسه (الصف {rec_y.row}) — "
                                            "قيد وعكسه يُلغي أحدهما الآخر، تحقق أنه إلغاء مقصود")

    # ---- الإجراء والأولوية ----
    def _finalize(self, r: MatchRow):
        f = r.flags
        if r.status == STATUS_MATCHED:
            r.priority = "—"
            r.action = "لا إجراء مطلوب"
            if "party" in f:
                r.priority = "منخفضة"
                r.action = "تأكد من أن اختلاف الاسم مقصود (مثلاً اسم الحساب مقابل المستفيد)"
        elif r.status == STATUS_DIFF:
            acts = []
            if "sign" in f:
                acts.append("راجع اتجاه العملية (مدين/دائن) في الملفين")
            if "currency" in f:
                acts.append("تحقق من عملة العملية في المستند الأصلي")
            if "fee" in f and r.explained:
                acts.append("اعتمد الفرق كرسوم/عمولة وسجّل قيد الرسوم إن لم يكن مسجلاً")
            elif "decimal" in f or "transpose" in f:
                acts.append("راجع المستند الأصلي — الخطأ على الأرجح في الإدخال لأحد الطرفين")
            elif "amount" in f:
                acts.append("راجع المستند الأصلي وصحّح المبلغ في الملف الخاطئ")
            if "date" in f:
                acts.append("راجع تاريخ القيد (قد تكون العملية مُرحّلة لفترة أخرى)")
            r.action = "؛ ".join(acts) or "راجع الفروقات"
            if r.explained:
                r.priority = "منخفضة"
            elif abs(r.diff) >= Decimal(1000) or "sign" in f or "currency" in f:
                r.priority = "عالية"
            else:
                r.priority = "متوسطة"
        elif r.status == STATUS_POSSIBLE:
            r.priority = "متوسطة"
            r.action = "راجع المطابقة المقترحة ثم اعتمدها أو ارفضها"
        elif r.status == STATUS_MISSING_B:
            r.priority = "عالية"
            r.action = ("اسأل الصرافة: هل نُفذت العملية؟ إن لم تُنفَّذ فصحّح قيد المحاسب، "
                        "وإن نُفذت فاطلب كشفاً مصحّحاً")
        elif r.status == STATUS_MISSING_A:
            r.priority = "عالية"
            r.action = "تحقق من مستند العملية ثم سجّلها في دفاتر المحاسب إن كانت سليمة"
        elif r.status == STATUS_DUP:
            r.priority = "عالية"
            r.action = "تحقق من التكرار؛ إن كان خطأً فاعكس القيد المكرر وصحّح الإجمالي"
        if "storno" in f and r.status in (STATUS_MISSING_A, STATUS_MISSING_B):
            r.priority = "منخفضة"
            r.action = "قيد وعكسه في الملف نفسه: صافي أثرهما صفر؛ تأكد فقط أن الإلغاء مقصود"

    # ---- الملخص والتحقق الحسابي ----
    def _summarize(self, rows, recs_a, recs_b, warnings, range_a, range_b) -> ReconResult:
        total_a: dict[str, Decimal] = defaultdict(Decimal)
        total_b: dict[str, Decimal] = defaultdict(Decimal)
        for r in recs_a:
            total_a[r.currency] += self.eff(r)
        for r in recs_b:
            total_b[r.currency] += self.eff(r)

        bridge: dict[str, dict[str, Decimal]] = defaultdict(lambda: defaultdict(Decimal))
        for row in rows:
            for cur, cat, amt in row.contribs:
                bridge[cur][cat] += amt
        integrity = {}
        for cur in set(total_a) | set(total_b) | set(bridge):
            actual = total_a.get(cur, Decimal(0)) - total_b.get(cur, Decimal(0))
            integrity[cur] = q2(sum(bridge[cur].values(), Decimal(0)) - actual)

        counts = {s: 0 for s in ALL_STATUSES}
        for r in rows:
            counts[r.status] += 1

        if any(v != 0 for v in integrity.values()):
            warnings.append("تنبيه: التحقق الحسابي لم يغلق على صفر — راجع القواعد أو أبلغ المطوّر.")

        return ReconResult(
            rows=rows, settings=self.S, rel=self.rel, rel_note=self.rel_note,
            count_a=len(recs_a), count_b=len(recs_b),
            total_a=dict(total_a), total_b=dict(total_b),
            bridge={c: dict(v) for c, v in bridge.items()},
            integrity=integrity, counts=counts, warnings=warnings,
            range_a=range_a, range_b=range_b)


def reconcile(recs_a: list[Record], recs_b: list[Record], settings: Optional[Settings] = None,
              progress: Optional[Callable[[int, str], None]] = None) -> ReconResult:
    return Engine(settings, progress).run(recs_a, recs_b)


# ---------------------------------------------------------------------------
# تعديلات المراجع بعد التشغيل
# ---------------------------------------------------------------------------
def set_decision(row: MatchRow, value: str) -> None:
    """يسجّل قرار المراجع. اعتماد عملية مرتبطة (مختلفة/محتملة) ينقلها إلى «متطابقة»،
    ومسح القرار أو رفضها يعيدها إلى حالتها الأصلية. (بعدها يُستدعى refresh_result)."""
    row.decision = value
    if value == "معتمد":
        if row.orig is None and row.a and row.b and row.status in (STATUS_DIFF, STATUS_POSSIBLE):
            row.orig = {"status": row.status, "priority": row.priority, "reasons": list(row.reasons),
                        "action": row.action}
            row.reasons = [f"اعتُمدت يدوياً من المراجع كمتطابقة (كانت: {row.status})"] + row.reasons
            row.status = STATUS_MATCHED
            row.priority = "—"
            row.action = "معتمدة من المراجع."
    elif row.orig is not None:
        o, row.orig = row.orig, None
        row.status, row.priority, row.reasons, row.action = o["status"], o["priority"], o["reasons"], o["action"]


def refresh_result(res: ReconResult) -> None:
    """يعيد حساب العدادات والجسر الحسابي والتحقق بعد أي تعديل على الصفوف."""
    bridge: dict[str, dict[str, Decimal]] = defaultdict(lambda: defaultdict(Decimal))
    counts = {s: 0 for s in ALL_STATUSES}
    for row in res.rows:
        counts[row.status] += 1
        for cur, cat, amt in row.contribs:
            bridge[cur][cat] += amt
    res.counts = counts
    res.bridge = {c: dict(v) for c, v in bridge.items()}
    res.integrity = {}
    for cur in set(res.total_a) | set(res.total_b) | set(bridge):
        actual = res.total_a.get(cur, Decimal(0)) - res.total_b.get(cur, Decimal(0))
        res.integrity[cur] = q2(sum(bridge[cur].values(), Decimal(0)) - actual)


def suggest_links(res: ReconResult, row: MatchRow) -> list[tuple[MatchRow, int, list[str]]]:
    """العمليات المقابلة الممكنة لعملية «مفقودة»، مرتبة بالأنسب أولاً: [(الصف، نسبة الملاءمة، الأسباب)]."""
    if row.status == STATUS_MISSING_B:
        pool = [r for r in res.rows if r.status == STATUS_MISSING_A]
        mine = row.a
    elif row.status == STATUS_MISSING_A:
        pool = [r for r in res.rows if r.status == STATUS_MISSING_B]
        mine = row.b
    else:
        return []
    eng = Engine(res.settings)
    eng.rel = res.rel
    recs_a = [r.a for r in res.rows if r.a]
    recs_b = [r.b for r in res.rows if r.b]
    eng._prepare_words(recs_a, recs_b)
    out = []
    for cand in pool:
        other = cand.b if row.status == STATUS_MISSING_B else cand.a
        why: list[str] = []
        score = 0.0
        am = abs(abs(mine.amount) - abs(other.amount))
        if am <= res.settings.amount_tol:
            score += 55
            why.append("نفس المبلغ")
        else:
            hi = max(abs(mine.amount), abs(other.amount)) or Decimal(1)
            p = float(am / hi)
            if p <= 0.05:
                score += 25 * (1 - p / 0.05)
                why.append(f"مبلغ قريب (فرق {fm_amt(am)})")
        if (mine.amount > 0) != (other.amount > 0) and mine.amount and other.amount:
            if (res.rel == -1):
                score += 3          # إشارات معكوسة متوقعة
            else:
                why.append("إشارة معاكسة")
        d = _days(mine.date, other.date)
        if d is not None:
            if d == 0:
                score += 20
                why.append("نفس التاريخ")
            elif d <= res.settings.fuzzy_days:
                score += 20 * (1 - d / (res.settings.fuzzy_days + 1))
                why.append(f"فارق {d} يوم")
        ov = eng._word_overlap(mine, other)
        if ov:
            score += 22 * min(ov, 1.0)
            why.append("أسماء/كلمات مشتركة")
        if mine.currency and other.currency and mine.currency != other.currency:
            score -= 15
            why.append("عملة مختلفة")
        out.append((cand, max(0, min(99, int(round(score)))), why))
    out.sort(key=lambda t: (-t[1], t[0].id))
    return out


def fm_amt(d: Decimal) -> str:
    return f"{q2(d):,.2f}"


def manual_link(res: ReconResult, row_a: MatchRow, row_b: MatchRow, approve: bool = True) -> MatchRow:
    """ربط يدوي بين صف «مفقود في الصرافة» وصف «مفقود في ملف المحاسب».
    approve=True: يُعتمد الربط فتنتقل العملية فوراً إلى «متطابقة»؛ False: تُترك معلّقة للمراجعة."""
    if not (row_a.status == STATUS_MISSING_B and row_a.a and row_b.status == STATUS_MISSING_A and row_b.b):
        raise ValueError("الربط اليدوي ممكن فقط بين عملية «مفقودة في الصرافة» وعملية «مفقودة في ملف المحاسب».")
    eng = Engine(res.settings)
    eng.rel = res.rel
    new = eng._row_from_pair(row_a.a, row_b.b, "ربط يدوي", 100)
    new.reasons = ["ربط يدوي بواسطة المراجع"] + [r for r in new.reasons if not r.startswith("تطابق تام")]
    eng._finalize(new)
    res.rows = [r for r in res.rows if r is not row_a and r is not row_b]
    new.id = max((r.id for r in res.rows), default=0) + 1
    res.rows.append(new)
    if approve:
        set_decision(new, "معتمد")
    refresh_result(res)
    return new


def amount_only_pairs(res: ReconResult) -> list[dict]:
    """أزواج «مفقودة في الصرافة» × «مفقودة في المحاسب» بمبلغ متساوٍ ولا يوجد ما يربطها (رقم/مرجع/اسم).
    تُعرض للمراجع فقط (ليست في التقرير حتى يعتمدها). كل عنصر: {a, b, score, why, unique, group}."""
    left = [r for r in res.rows if r.status == STATUS_MISSING_B and r.a]
    right = [r for r in res.rows if r.status == STATUS_MISSING_A and r.b]
    tol = res.settings.amount_tol
    groups: dict = {}
    for r in left:
        groups.setdefault(q2(abs(r.a.amount)), ([], []))[0].append(r)
    for r in right:
        groups.setdefault(q2(abs(r.b.amount)), ([], []))[1].append(r)
    out: list[dict] = []
    for amt, (ls, rs) in groups.items():
        if not ls or not rs or amt == 0:
            continue
        # تأكد من توافق العملة
        used: set = set()
        cands = []
        for la in ls:
            sg = {id(x): (s, w) for x, s, w in suggest_links(res, la)}
            for rb in rs:
                cur_a, cur_b = la.a.currency, rb.b.currency
                if cur_a and cur_b and cur_a != cur_b:
                    continue
                s, w = sg.get(id(rb), (0, []))
                cands.append((s, la, rb, w))
        cands.sort(key=lambda t: (-t[0], t[1].id, t[2].id))
        pairs = []
        for s, la, rb, w in cands:
            if la.id in used or rb.id in {p[2].id for p in pairs}:
                continue
            used.add(la.id)
            pairs.append((s, la, rb, w))
        n_l = len({la.id for _s, la, _r, _w in cands})
        n_r = len({rb.id for _s, _l, rb, _w in cands})
        for s, la, rb, w in pairs:
            out.append({"a": la, "b": rb, "score": s, "why": w, "unique": n_l == 1 and n_r == 1,
                        "group": f"{n_l}×{n_r}"})
    out.sort(key=lambda d: (not d["unique"], -d["score"], d["a"].id))
    return out
