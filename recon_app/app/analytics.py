"""تحليل نتائج المطابقة: أين المشكلة؟ من الأطراف الأكثر تكراراً؟ ما الأنماط المتكررة؟

دوال نقية (بدون واجهة) تُنتج `Analysis` تستهلكها الواجهة وتقرير Excel معاً.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal
from typing import Optional

from .engine import (CAT_DUP_A, CAT_DUP_B, CAT_FEE, CAT_MISS_A, CAT_MISS_B, STATUS_DIFF, STATUS_DUP,
                     STATUS_MATCHED, STATUS_MISSING_A, STATUS_MISSING_B, STATUS_POSSIBLE, MatchRow,
                     ReconResult)
from .normalize import q2

ISSUE_STATUSES = (STATUS_DIFF, STATUS_POSSIBLE, STATUS_MISSING_A, STATUS_MISSING_B, STATUS_DUP)
D0 = Decimal(0)


@dataclass
class Finding:
    level: str            # high | mid | info | good
    title: str
    detail: str
    action: Optional[dict] = None     # {"status": ..} أو {"search": ..} لفتح العمليات المعنية
    action_label: str = "عرض العمليات"


@dataclass
class CatStat:
    category: str
    currency: str
    count: int
    amount: Decimal       # مجموع القيم المطلقة


@dataclass
class PartyStat:
    party: str
    total: int            # كل عملياته
    issues: int           # العمليات غير المتطابقة
    exposure: dict        # عملة -> مبلغ مطلق غير مطابق

    @property
    def rate(self) -> float:
        return self.issues / self.total if self.total else 0.0


@dataclass
class Analysis:
    rows_total: int = 0
    issues_total: int = 0
    exposure: dict = field(default_factory=dict)         # عملة -> {"abs":, "net":, "n":}
    categories: list = field(default_factory=list)       # [CatStat] مرتبة بالمبلغ
    parties: list = field(default_factory=list)          # [PartyStat] مرتبة
    timeline: list = field(default_factory=list)         # [(تسمية, إجمالي, مشاكل)]
    timeline_unit: str = "يوم"
    patterns: list = field(default_factory=list)         # [(وصف, عدد, عملة)]
    pareto: Optional[dict] = None
    findings: list = field(default_factory=list)         # [Finding]
    has_dates: bool = False
    has_parties: bool = False


def _row_party(r: MatchRow) -> str:
    for rec in (r.a, r.b):
        if rec and rec.party:
            return " ".join(rec.party.split())
    return ""


def _row_date(r: MatchRow) -> Optional[date]:
    for rec in (r.a, r.b):
        if rec and rec.date:
            return rec.date
    return None


def _is_issue(r: MatchRow) -> bool:
    return r.status in ISSUE_STATUSES


def _row_exposure(r: MatchRow) -> dict:
    """القيمة المطلقة غير المطابقة لكل عملة في هذه العملية."""
    out: dict[str, Decimal] = defaultdict(lambda: D0)
    if r.contribs:
        for cur, _cat, amt in r.contribs:
            out[cur or r.currency] += abs(amt)
    else:
        v = abs(r.diff) if r.diff else abs((r.amt_a or r.amt_b or D0))
        out[r.currency] += v
    return out


def _fmt(d: Decimal) -> str:
    return f"{q2(d):,.2f}"


def _fmt_cur(d: dict) -> str:
    return " + ".join(f"{_fmt(v)} {c}".strip() for c, v in sorted(d.items(), key=lambda kv: -kv[1]) if v) or "—"


def analyze(res: ReconResult) -> Analysis:
    an = Analysis(rows_total=len(res.rows))
    issues = [r for r in res.rows if _is_issue(r)]
    an.issues_total = len(issues)

    # ---- الانكشاف المالي لكل عملة + التصنيفات
    exp: dict[str, dict] = {}
    cats: dict[tuple, list] = {}
    for r in issues:
        for cur, cat, amt in (r.contribs or [(r.currency, r.status, r.diff)]):
            cur = cur or r.currency
            e = exp.setdefault(cur, {"abs": D0, "net": D0, "n": 0})
            e["abs"] += abs(amt)
            e["net"] += amt
            c = cats.setdefault((cat, cur), [0, D0])
            c[0] += 1
            c[1] += abs(amt)
        for cur in {(c or r.currency) for c, _x, _y in (r.contribs or [(r.currency, "", D0)])}:
            exp.setdefault(cur, {"abs": D0, "net": D0, "n": 0})["n"] += 1
    an.exposure = exp
    an.categories = sorted((CatStat(k[0], k[1], v[0], v[1]) for k, v in cats.items()),
                           key=lambda c: -c.amount)

    # ---- الأطراف
    pt: dict[str, list] = {}
    for r in res.rows:
        p = _row_party(r)
        if not p:
            continue
        s = pt.setdefault(p, [0, 0, defaultdict(lambda: D0)])
        s[0] += 1
        if _is_issue(r):
            s[1] += 1
            for cur, v in _row_exposure(r).items():
                s[2][cur] += v
    an.has_parties = bool(pt)
    an.parties = sorted((PartyStat(p, v[0], v[1], dict(v[2])) for p, v in pt.items() if v[1]),
                        key=lambda s: (-s.issues, -sum(s.exposure.values())))[:12]

    # ---- المحور الزمني
    dated = [(r, _row_date(r)) for r in res.rows]
    dated = [(r, d) for r, d in dated if d]
    an.has_dates = bool(dated)
    if dated:
        lo, hi = min(d for _, d in dated), max(d for _, d in dated)
        span = (hi - lo).days + 1
        weekly = span > 45
        monthly = span > 400
        an.timeline_unit = "شهر" if monthly else ("أسبوع" if weekly else "يوم")

        def bucket(d: date):
            if monthly:
                return date(d.year, d.month, 1)
            if weekly:
                return d - timedelta(days=d.weekday())
            return d
        agg: dict = {}
        for r, d in dated:
            b = agg.setdefault(bucket(d), [0, 0])
            b[0] += 1
            b[1] += 1 if _is_issue(r) else 0
        an.timeline = [(k.strftime("%Y-%m" if monthly else "%m-%d"), v[0], v[1])
                       for k, v in sorted(agg.items())]

    # ---- أنماط متكررة
    diffs = Counter()
    for r in res.rows:
        if r.status == STATUS_DIFF and r.diff and abs(r.diff) > Decimal("0.009"):
            diffs[(q2(abs(r.diff)), r.currency)] += 1
    an.patterns = [(f"فرق ثابت قدره {_fmt(v)}", n, cur) for (v, cur), n in diffs.most_common(5) if n >= 3]

    # ---- باريتو: كم عملية تشكّل معظم القيمة؟
    if issues:
        main_cur = max(exp, key=lambda c: exp[c]["abs"]) if exp else ""
        vals = sorted((_row_exposure(r).get(main_cur, D0) for r in issues), reverse=True)
        vals = [v for v in vals if v > 0]
        tot = sum(vals, D0)
        if tot > 0 and len(vals) >= 5:
            acc, k = D0, 0
            for v in vals:
                acc += v
                k += 1
                if acc >= tot * Decimal("0.8"):
                    break
            an.pareto = {"cur": main_cur, "k": k, "n": len(vals), "share": 0.8}

    an.findings = _findings(res, an, issues)
    return an


def _findings(res: ReconResult, an: Analysis, issues: list) -> list[Finding]:
    out: list[Finding] = []
    n = max(len(res.rows), 1)
    rate = res.counts.get(STATUS_MATCHED, 0) / n

    if not issues:
        out.append(Finding("good", "كل شيء متطابق", "لا توجد عمليات تحتاج متابعة في هذه المطابقة."))
        return out

    # 1) أكبر مصدر للفرق المالي
    if an.categories:
        top = an.categories[0]
        tot = sum((c.amount for c in an.categories if c.currency == top.currency), D0) or Decimal(1)
        share = float(top.amount / tot * 100)
        status_map = {CAT_MISS_B: STATUS_MISSING_B, CAT_MISS_A: STATUS_MISSING_A,
                      CAT_DUP_A: STATUS_DUP, CAT_DUP_B: STATUS_DUP}
        act = {"status": status_map[top.category]} if top.category in status_map else {"status": STATUS_DIFF}
        out.append(Finding(
            "high" if share >= 50 else "mid",
            f"أكبر مصدر للفرق: {top.category}",
            f"يمثّل {share:.0f}٪ من القيمة غير المطابقة بعملة {top.currency or '—'} "
            f"({_fmt(top.amount)} في {top.count} عملية). ابدأ به لتسوية أكبر جزء بأقل جهد.", act))

    # 2) تركّز القيمة (باريتو)
    if an.pareto:
        p = an.pareto
        out.append(Finding(
            "info", f"{p['k']} عملية فقط تشكّل 80٪ من القيمة",
            f"من أصل {p['n']} عملية غير مطابقة بعملة {p['cur'] or '—'}؛ راجعها أولاً (مرتبة بالأهم في وضع المراجعة السريعة).",
            {"priority": "عالية"}, "عرض الأعلى أولوية"))

    # 3) أطراف متكررة
    if an.parties:
        p = an.parties[0]
        if p.issues >= 3 and p.rate >= 0.3:
            out.append(Finding(
                "mid", f"الطرف «{p.party}» متكرر المشاكل",
                f"{p.issues} من {p.total} عملية لديه غير مطابقة ({p.rate * 100:.0f}٪)، بقيمة {_fmt_cur(p.exposure)}. "
                "قد يكون السبب كتابة اسم/مرجع مختلفة أو مشكلة منهجية معه.", {"search": p.party}, "عرض عملياته"))

    # 4) فرق ثابت متكرر
    if an.patterns:
        d, c, cur = an.patterns[0]
        out.append(Finding(
            "mid", f"{c} عمليات بنفس الفرق ({d.replace('فرق ثابت قدره ', '')} {cur})".strip(),
            "تكرار فرق ثابت بالضبط يدل غالباً على رسوم/عمولة غير مسجلة في أحد الملفين؛ اعتمادها دفعة واحدة يوفر وقتاً.",
            {"status": STATUS_DIFF}))

    # 4ب) نفس الرقم لكن بإشارة معكوسة (مدين/دائن): احتياط خاص
    sgn = [r for r in res.rows if "sign" in r.flags and r.a and r.b]
    if sgn:
        cur_eff: dict[str, Decimal] = defaultdict(lambda: D0)
        for r in sgn:
            cur_eff[r.currency] += abs(r.diff)
        out.append(Finding(
            "high", f"{len(sgn)} عملية بنفس الرقم لكن بإشارة معكوسة (مدين/دائن)",
            f"أثرها الفعلي على الرصيد ضعف المبلغ ({_fmt_cur(dict(cur_eff))}). غالباً خطأ في اتجاه القيد لا في الرقم؛ "
            "صحّح الاتجاه في أحد الملفين ثم أعد المطابقة. لم تُعتمد تلقائياً كمتطابقة عن قصد.",
            {"search": "معكوسة"}, "عرض العمليات"))
    pa, pb = res.rel_note or "", ""
    if "معكوسة مقابل" in pa or "مقابل" in pa:
        import re as _re
        m = _re.search(r"\((\d+) عملية معكوسة مقابل (\d+)\)|\((\d+) مقابل (\d+) معكوسة\)", pa)
        if m:
            g = [int(x) for x in m.groups() if x]
            lo = min(g)
            if lo >= 1:
                out.append(Finding("info", "اتجاه الإشارات مختلط بين الملفين",
                                   f"{pa}. الاتجاه الغالب اعتُمد للمقارنة والأقلية فُصلت للمراجعة؛ "
                                   "إن كان الاختلاف مقصوداً غيّر «اتجاه الإشارات» يدوياً من صفحة القواعد."))

    # 5) عمليات مفقودة في الطرفين بنفس المبلغ (مع التمييز بين نفس الإشارة وإشارة معكوسة)
    ma = [r for r in res.rows if r.status == STATUS_MISSING_A]
    mb = [r for r in res.rows if r.status == STATUS_MISSING_B]
    sig_abs = Counter((q2(abs(r.amt_b or D0)), r.currency) for r in ma)
    sig_signed = Counter((q2(r.amt_b or D0), r.currency) for r in ma)
    pairs_same = pairs_opp = 0
    for r in mb:
        v = q2(r.amt_a or D0)
        if sig_signed.get((v, r.currency)):
            pairs_same += 1
        elif sig_abs.get((abs(v), r.currency)):
            pairs_opp += 1
    if pairs_same:
        out.append(Finding(
            "mid", f"{pairs_same} عملية مفقودة في طرف ولها نظير بنفس المبلغ في الطرف الآخر",
            "غالباً هي نفس العملية بمرجع أو بيان مختلف. استخدم «ربط يدوي بين المحدّدتين» لدمجها بدل معاملتها كمفقودة.",
            {"status": STATUS_MISSING_B}, "عرض المفقودات"))
    if pairs_opp:
        out.append(Finding(
            "mid", f"{pairs_opp} عملية مفقودة لها نظير بنفس الرقم لكن بإشارة معاكسة",
            "قد تكون نفس العملية مسجّلة بالاتجاه المعاكس (مدين مقابل دائن)، وقد تكون عملية عكسية/إلغاء حقيقية. "
            "تحقق من البيان والتاريخ قبل الربط اليدوي.", {"status": STATUS_MISSING_B}, "عرض المفقودات"))

    # 6) مكررات
    dup = res.counts.get(STATUS_DUP, 0)
    if dup:
        out.append(Finding("mid", f"{dup} عملية مكررة", "تكرار نفس العملية يضخّم أحد الجانبين؛ تحقق أنه ليس تسجيلاً مزدوجاً.",
                           {"status": STATUS_DUP}))

    # 7) نقطة زمنية شاذة
    if len(an.timeline) >= 4:
        vals = [t[2] for t in an.timeline]
        avg = sum(vals) / len(vals)
        i = max(range(len(vals)), key=lambda k: vals[k])
        if vals[i] >= 3 and vals[i] >= 2.5 * max(avg, 0.5):
            out.append(Finding("info", f"ذروة مشاكل عند {an.timeline[i][0]}",
                               f"{vals[i]} عملية غير مطابقة في {an.timeline_unit} واحد مقابل متوسط {avg:.1f}؛ "
                               "قد يشير إلى دفعة مستوردة بشكل خاطئ أو حدث محدد في ذلك التاريخ."))

    # 8) صحة عامة
    if rate >= 0.95:
        out.append(Finding("good", f"جودة عالية: {rate * 100:.1f}٪ متطابقة", "ما تبقّى قليل ويمكن إنهاؤه بمراجعة سريعة."))
    elif rate < 0.6:
        out.append(Finding("high", f"نسبة التطابق منخفضة ({rate * 100:.0f}٪)",
                           "قد تكون الأعمدة المختارة للمطابقة غير مناسبة لهذين الملفين؛ راجع «أعمدة المطابقة بالأولوية» في صفحة القواعد."))
    return out
