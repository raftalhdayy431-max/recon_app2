"""قواعد مفاتيح المطابقة الديناميكية.

كل «قاعدة» = مجموعة أعمدة من ملف المحاسب + مجموعة أعمدة من ملف الصرافة + نمط المطابقة.
تُطبَّق القواعد بالترتيب (الأولوية)، وما لم يُطابَق بقاعدة ينتقل إلى التي تليها.

أنماط المطابقة:
  text  : استخراج الأرقام/الرموز (مثل رقم الحوالة) من نص العمود (كالبيان) ومطابقتها بين الملفين.
  exact : مطابقة قيمة العمود كاملة (بعد التطبيع) كمرجع.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Optional

from .normalize import clean_text, norm_ref

MODE_TEXT = "text"
MODE_EXACT = "exact"
MODE_LABELS = {MODE_TEXT: "رقم/رمز داخل النص", MODE_EXACT: "قيمة العمود كاملة"}

_DATE_LIKE = re.compile(r"(?:\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01]))")
_TOKEN = re.compile(r"[A-Za-z0-9][A-Za-z0-9\-_/]*[A-Za-z0-9]|[A-Za-z0-9]")
# أدوار لا تُستعمل كمفاتيح (قيم مالية/تاريخية/وصفية عامة)
_NON_KEY_ROLES = {"amount", "debit", "credit", "date", "currency", "fee", "party"}


@dataclass
class KeyRule:
    name: str
    a: list[str] = field(default_factory=list)
    b: list[str] = field(default_factory=list)
    mode: str = MODE_TEXT
    enabled: bool = True

    def to_dict(self) -> dict:
        return {"name": self.name, "a": list(self.a), "b": list(self.b), "mode": self.mode, "enabled": self.enabled}

    @staticmethod
    def from_dict(d: dict) -> "KeyRule":
        return KeyRule(name=str(d.get("name", "")), a=list(d.get("a", [])), b=list(d.get("b", [])),
                       mode=d.get("mode", MODE_TEXT) if d.get("mode") in MODE_LABELS else MODE_TEXT,
                       enabled=bool(d.get("enabled", True)))


def other_columns(columns: list[str], roles: dict[str, str]) -> list[str]:
    """الأعمدة غير المُسنَدة لدور مالي/تاريخي (مرشّحة كمفاتيح احتياطية)."""
    used = {c for r, c in roles.items() if r in _NON_KEY_ROLES or r in ("desc", "ref", "ref2")}
    return [c for c in columns if c not in used and not re.fullmatch(r"عمود \d+", c)]


def default_rules(cols_a: list[str], roles_a: dict[str, str], cols_b: list[str],
                  roles_b: dict[str, str]) -> list[KeyRule]:
    """السلسلة التلقائية: البيان أولاً، ثم المرجع، ثم التقاطعات، ثم بقية الأعمدة."""
    g = lambda roles, *rs: [roles[r] for r in rs if roles.get(r)]
    rules: list[KeyRule] = []

    def add(name, a, b, mode=MODE_TEXT):
        if a and b:
            rules.append(KeyRule(name, a, b, mode))

    add("البيان", g(roles_a, "desc"), g(roles_b, "desc"))
    add("المرجع", g(roles_a, "ref"), g(roles_b, "ref"), MODE_EXACT)
    add("بيان المحاسب ↔ مرجع الصرافة", g(roles_a, "desc", "ref2"), g(roles_b, "ref", "ref2"))
    add("مرجع المحاسب ↔ بيان الصرافة", g(roles_a, "ref", "ref2"), g(roles_b, "desc", "ref2"))
    add("السند / المرجع الثانوي", g(roles_a, "ref2"), g(roles_b, "ref2"), MODE_EXACT)
    oa, ob = other_columns(cols_a, roles_a), other_columns(cols_b, roles_b)
    add("بقية الأعمدة", oa, ob)
    # إزالة تكرار القواعد المتطابقة
    seen, out = set(), []
    for r in rules:
        k = (tuple(r.a), tuple(r.b), r.mode)
        if k not in seen:
            seen.add(k)
            out.append(r)
    return out


def extract_tokens(value: Any, mode: str, min_len: int) -> list[str]:
    """القيم المرشّحة للمطابقة من محتوى خلية واحدة."""
    if value is None or value == "":
        return []
    if mode == MODE_EXACT:
        n = norm_ref(value)
        return [n] if n else []
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = clean_text(value) if isinstance(value, str) else str(value)
    out, seen = [], set()
    for t in _TOKEN.findall(text):
        if _DATE_LIKE.fullmatch(t):
            continue          # التواريخ ليست مفاتيح مطابقة
        n = norm_ref(t)
        if re.fullmatch(r"\d{1,4}20[2-3]\d", n) and len(n) <= 8:
            continue          # «11-2026» / «10/2026م» بعد إزالة الفواصل = شهر وسنة لا رقم عملية
        if len(n) >= min_len and n not in seen:
            seen.add(n)
            out.append(n)
    return out
