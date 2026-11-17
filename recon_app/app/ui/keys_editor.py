"""محرر «أعمدة المطابقة بالأولوية»: يحدد المستخدم أي أعمدة تُعتمد للمطابقة الأولية وما بعدها."""
from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (QCheckBox, QComboBox, QFrame, QHBoxLayout, QLabel, QLineEdit, QMenu,
                               QPushButton, QVBoxLayout, QWidget)

from ..keys import MODE_LABELS, KeyRule, default_rules
from .widgets import IconTile, icon


class _ColPicker(QPushButton):
    """زر يفتح قائمة أعمدة بمربعات اختيار (متعدد)."""
    changed = Signal()

    def __init__(self, title: str):
        super().__init__()
        self.title = title
        self.setObjectName("Ghost")
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumWidth(190)
        self.selected: list[str] = []
        self.options: list[str] = []
        self.setStyleSheet("text-align:right; padding:8px 12px;")
        self._refresh()

    def set_state(self, selected: list[str], options: list[str]):
        self.selected = list(selected)
        self.options = list(dict.fromkeys(list(options) + list(selected)))
        self._refresh()

    def _refresh(self):
        n = len(self.selected)
        txt = "— اختر أعمدة —" if n == 0 else (self.selected[0] if n == 1 else f"{self.selected[0]}  +{n - 1}")
        self.setText(f"{self.title}:  {txt}")
        self.setToolTip("\n".join(f"{i + 1}. {c}" for i, c in enumerate(self.selected)) or "لم تُحدَّد أعمدة")

    def mousePressEvent(self, e):
        menu = QMenu(self)
        if not self.options:
            a = menu.addAction("حمّل الملف أولاً لتظهر أعمدته")
            a.setEnabled(False)
        for c in self.options:
            act = QAction(c, menu, checkable=True)
            act.setChecked(c in self.selected)
            act.toggled.connect(lambda on, col=c: self._toggle(col, on))
            menu.addAction(act)
        menu.exec(self.mapToGlobal(self.rect().bottomLeft()))

    def _toggle(self, col: str, on: bool):
        if on and col not in self.selected:
            self.selected.append(col)
        elif not on and col in self.selected:
            self.selected.remove(col)
        self._refresh()
        self.changed.emit()


class KeyRulesEditor(QFrame):
    changed = Signal()

    def __init__(self):
        super().__init__()
        self.setObjectName("Card")
        self.cols_a: list[str] = []
        self.cols_b: list[str] = []
        self.roles_a: dict = {}
        self.roles_b: dict = {}
        self.auto = True
        self.model: list[KeyRule] = []
        self._building = False

        v = QVBoxLayout(self)
        v.setContentsMargins(30, 26, 30, 26)
        v.setSpacing(14)
        head = QHBoxLayout()
        head.setSpacing(12)
        head.addWidget(IconTile("link", "#0D9488", 44))
        tv = QVBoxLayout()
        tv.setSpacing(2)
        t = QLabel("أعمدة المطابقة بالأولوية")
        t.setObjectName("H2")
        s = QLabel("يطابق البرنامج بالقاعدة الأولى؛ وما لم يُطابَق ينتقل إلى التي تليها، وهكذا. "
                   "الأعمدة تُسحب تلقائياً من الملفين، ويمكنك تعديل الترتيب والأعمدة.")
        s.setObjectName("Muted")
        s.setWordWrap(True)
        tv.addWidget(t)
        tv.addWidget(s)
        head.addLayout(tv, 1)
        self.badge = QLabel("تلقائي")
        self.badge.setObjectName("BadgeOk")
        head.addWidget(self.badge)
        v.addLayout(head)

        bar = QHBoxLayout()
        self.btn_add = QPushButton("+  إضافة قاعدة")
        self.btn_add.setObjectName("Ghost")
        self.btn_add.setCursor(Qt.PointingHandCursor)
        self.btn_add.clicked.connect(self._add)
        self.btn_reset = QPushButton("إعادة الضبط التلقائي")
        self.btn_reset.setObjectName("Ghost")
        self.btn_reset.setCursor(Qt.PointingHandCursor)
        self.btn_reset.clicked.connect(self.reset_auto)
        bar.addWidget(self.btn_add)
        bar.addWidget(self.btn_reset)
        bar.addStretch(1)
        v.addLayout(bar)

        self.hint = QLabel("")
        self.hint.setObjectName("Warn")
        self.hint.setWordWrap(True)
        self.hint.hide()
        v.addWidget(self.hint)
        self.box = QVBoxLayout()
        self.box.setSpacing(10)
        v.addLayout(self.box)
        self.legend = QLabel("«رقم/رمز داخل النص»: يستخرج الأرقام والرموز (مثل رقم الحوالة) من نص العمود كالبيان ويطابقها "
                             "بين الملفين.   «قيمة العمود كاملة»: يطابق محتوى العمود كمرجع.")
        self.legend.setObjectName("Muted")
        self.legend.setWordWrap(True)
        v.addWidget(self.legend)
        self._rebuild()

    # ---------------------------------------------------------------- واجهة عامة
    def set_columns(self, cols_a, roles_a, cols_b, roles_b):
        self.cols_a, self.cols_b = list(cols_a or []), list(cols_b or [])
        self.roles_a = {k: v for k, v in (roles_a or {}).items() if v}
        self.roles_b = {k: v for k, v in (roles_b or {}).items() if v}
        if self.auto:
            self.model = default_rules(self.cols_a, self.roles_a, self.cols_b, self.roles_b)
        self._rebuild()

    def set_rules(self, rules: Optional[list]):
        if rules:
            self.auto = False
            self.model = [KeyRule.from_dict(d) for d in rules]
        else:
            self.auto = True
            self.model = default_rules(self.cols_a, self.roles_a, self.cols_b, self.roles_b)
        self._rebuild()

    def rules(self) -> Optional[list]:
        """None = تلقائي (يُشتق من أعمدة كل ملف عند التشغيل)، وإلا قائمة القواعد المخصصة."""
        return None if self.auto else [r.to_dict() for r in self.model]

    # ---------------------------------------------------------------- داخلي
    def _touch(self):
        if not self._building:
            self.auto = False
            self._update_badge()
            self.changed.emit()

    def _update_badge(self):
        self.badge.setText("تلقائي" if self.auto else "مخصص")
        self.badge.setObjectName("BadgeOk" if self.auto else "Badge")
        self.badge.style().unpolish(self.badge)
        self.badge.style().polish(self.badge)

    def reset_auto(self):
        self.auto = True
        self.model = default_rules(self.cols_a, self.roles_a, self.cols_b, self.roles_b)
        self._rebuild()
        self.changed.emit()

    def _add(self):
        self.model.append(KeyRule(f"قاعدة {len(self.model) + 1}", [], []))
        self._touch()
        self._rebuild()

    def _move(self, i: int, d: int):
        j = i + d
        if 0 <= j < len(self.model):
            self.model[i], self.model[j] = self.model[j], self.model[i]
            self._touch()
            self._rebuild()

    def _delete(self, i: int):
        del self.model[i]
        self._touch()
        self._rebuild()

    def _rebuild(self):
        self._building = True
        while self.box.count():
            it = self.box.takeAt(0)
            ow = it.widget()
            if ow is not None:
                ow.hide()
                ow.setParent(None)
                ow.deleteLater()
        self._update_badge()
        no_files = not (self.cols_a or self.cols_b)
        if no_files and self.auto:
            self.hint.setText("حمّل ملفّي المحاسب والصرافة في الخطوة الأولى لتظهر هنا القواعد المقترحة وأعمدة الملفين.")
            self.hint.show()
        elif not self.model:
            self.hint.setText("لا توجد أعمدة مشتركة صالحة للمطابقة. أضف قاعدة يدوياً، أو ستُستخدم المطابقة الذكية "
                              "(مبلغ/تاريخ/اسم) فقط.")
            self.hint.show()
        else:
            self.hint.hide()
        for i, r in enumerate(self.model):
            w = self._row(i, r)
            self.box.addWidget(w)
            w.show()
        self._building = False

    def _row(self, i: int, r: KeyRule) -> QWidget:
        f = QFrame()
        f.setObjectName("CardFlat")
        h = QHBoxLayout(f)
        h.setContentsMargins(14, 10, 12, 10)
        h.setSpacing(10)
        num = QLabel(str(i + 1))
        num.setFixedSize(30, 30)
        num.setAlignment(Qt.AlignCenter)
        num.setStyleSheet("background:#0D9488; color:white; border-radius:15px; font-weight:800;")
        h.addWidget(num)
        en = QCheckBox()
        en.setChecked(r.enabled)
        en.setToolTip("تفعيل/تعطيل هذه القاعدة")
        en.toggled.connect(lambda on, k=r: (setattr(k, "enabled", on), self._touch()))
        h.addWidget(en)
        name = QLineEdit(r.name)
        name.setMinimumWidth(150)
        name.setPlaceholderText("اسم القاعدة")
        name.textEdited.connect(lambda t, k=r: (setattr(k, "name", t), self._touch()))
        h.addWidget(name, 1)
        mode = QComboBox()
        for k, lbl in MODE_LABELS.items():
            mode.addItem(lbl, k)
        mode.setCurrentIndex(max(mode.findData(r.mode), 0))
        mode.currentIndexChanged.connect(lambda _i, k=r, m=mode: (setattr(k, "mode", m.currentData()), self._touch()))
        h.addWidget(mode)
        for side, title, opts in (("a", "المحاسب", self.cols_a), ("b", "الصرافة", self.cols_b)):
            p = _ColPicker(title)
            p.set_state(getattr(r, side), opts)
            p.changed.connect(lambda k=r, s=side, pk=p: (setattr(k, s, list(pk.selected)), self._touch()))
            h.addWidget(p)
        for txt, tip, fn in (("▲", "رفع الأولوية", lambda _=False, k=i: self._move(k, -1)),
                             ("▼", "خفض الأولوية", lambda _=False, k=i: self._move(k, 1)),
                             ("✕", "حذف القاعدة", lambda _=False, k=i: self._delete(k))):
            b = QPushButton(txt)
            b.setObjectName("Ghost")
            b.setToolTip(tip)
            b.setFixedWidth(38)
            b.setStyleSheet("padding:8px 0px;")
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(fn)
            h.addWidget(b)
        return f
