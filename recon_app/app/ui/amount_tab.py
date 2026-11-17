"""تبويب «تطابق بالمبلغ فقط»: عمليات متساوية المبلغ بلا رقم/مرجع/اسم يربطها. للاطلاع، ولا تدخل التقرير إلا باعتماد المراجع."""
from __future__ import annotations

from decimal import Decimal

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QFrame, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
                               QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from ..engine import amount_only_pairs, manual_link
from ..normalize import q2


def _short(rec, n=95) -> str:
    return " ".join((rec.desc or "").split())[:n] or "—"


class AmountPairsView(QWidget):
    changed = Signal(int)            # عدد الأزواج التي اعتُمدت

    COLS = ["", "الوضوح", "المبلغ", "التاريخ", "بيان المحاسب (#)", "بيان الصرافة (#)", "ملاحظة"]

    def __init__(self):
        super().__init__()
        self.setObjectName("Page")
        self.res = None
        self.pairs: list[dict] = []
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 16, 0, 0)
        v.setSpacing(10)

        info = QLabel("عمليات مفقودة في الجهتين <b>مبلغها متساوٍ</b> ولا يوجد رقم حوالة أو مرجع أو اسم يؤكد أنها واحدة. "
                      "لا تظهر في تقرير Excel كمتطابقة إلا إذا اعتمدتها أنت — راجعها ثم حدّد ما تراه صحيحاً.")
        info.setObjectName("Warn")
        info.setWordWrap(True)
        info.setTextFormat(Qt.RichText)
        v.addWidget(info)

        bar = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("🔍  ابحث بالمبلغ أو أي كلمة…")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self._filter)
        bar.addWidget(self.search, 1)
        for text, fn in (("تحديد الواضحة", self._sel_unique), ("تحديد الكل", lambda: self._set_all(True)),
                         ("إلغاء التحديد", lambda: self._set_all(False))):
            b = QPushButton(text)
            b.setObjectName("Chip")
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(fn)
            bar.addWidget(b)
        v.addLayout(bar)

        self.tbl = QTableWidget(0, len(self.COLS))
        self.tbl.setHorizontalHeaderLabels(self.COLS)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setAlternatingRowColors(True)
        self.tbl.setWordWrap(True)
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(4, QHeaderView.Stretch)
        hh.setSectionResizeMode(5, QHeaderView.Stretch)
        for c, w in ((0, 44), (1, 100), (2, 130), (3, 120), (6, 260)):
            self.tbl.setColumnWidth(c, w)
        v.addWidget(self.tbl, 1)

        foot = QHBoxLayout()
        self.lbl = QLabel("")
        self.lbl.setObjectName("Muted")
        foot.addWidget(self.lbl, 1)
        self.btn = QPushButton("")
        self.btn.setObjectName("Good")
        self.btn.setCursor(Qt.PointingHandCursor)
        self.btn.clicked.connect(self.approve_selected)
        foot.addWidget(self.btn)
        v.addLayout(foot)
        self.checks: list[QCheckBox] = []

    # ------------------------------------------------------------------
    def set_result(self, res):
        self.res = res
        self.pairs = amount_only_pairs(res) if res else []
        self.tbl.setRowCount(len(self.pairs))
        self.checks = []
        for i, p in enumerate(self.pairs):
            a, b = p["a"], p["b"]
            amt = q2(abs(a.a.amount))
            cur = a.a.currency or b.b.currency
            d = a.a.date or b.b.date
            note = ("المبلغ فريد في الجهتين" if p["unique"] else f"مبلغ مكرر ({p['group']}) — تأكد من التاريخ والاسم")
            if p["why"]:
                note += "  ·  " + "، ".join(w for w in p["why"] if w != "نفس المبلغ")
            vals = ["", "واضح" if p["unique"] else "يحتاج تدقيق", f"{amt:,.2f} {cur}".strip(),
                    d.strftime("%Y-%m-%d") if d else "—",
                    f"({a.id})  {_short(a.a)}", f"({b.id})  {_short(b.b)}", note]
            for j, val in enumerate(vals):
                it = QTableWidgetItem(val)
                it.setTextAlignment(Qt.AlignCenter if j in (1, 2, 3) else Qt.AlignVCenter | Qt.AlignLeading)
                if j == 1:
                    it.setForeground(QColor("#1F9D55" if p["unique"] else "#D99100"))
                    f = it.font()
                    f.setBold(True)
                    it.setFont(f)
                self.tbl.setItem(i, j, it)
            cb = QCheckBox()
            cb.toggled.connect(self._count)
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setAlignment(Qt.AlignCenter)
            h.addWidget(cb)
            self.tbl.setCellWidget(i, 0, w)
            self.checks.append(cb)
            self.tbl.setRowHeight(i, 46)
        self._count()

    def count(self) -> int:
        return len(self.pairs)

    def _count(self, *_):
        n = sum(1 for c in self.checks if c.isChecked())
        self.btn.setEnabled(n > 0)
        self.btn.setText(f"✔  اعتماد المحدد ({n}) وإضافته إلى المتطابقات" if n else "حدّد ما تريد اعتماده")
        if self.pairs:
            u = sum(1 for p in self.pairs if p["unique"])
            self.lbl.setText(f"{len(self.pairs)} زوج مقترح · {u} واضح (مبلغ فريد) · {len(self.pairs) - u} يحتاج تدقيق")
        else:
            self.lbl.setText("لا توجد عمليات متساوية المبلغ بين المفقودات — كل شيء إما مطابق أو مختلف المبلغ.")

    def _set_all(self, on: bool):
        for i, c in enumerate(self.checks):
            if not self.tbl.isRowHidden(i):
                c.setChecked(on)

    def _sel_unique(self):
        for i, c in enumerate(self.checks):
            c.setChecked(self.pairs[i]["unique"] and not self.tbl.isRowHidden(i))

    def _filter(self, text: str):
        t = text.strip().lower().replace(",", "")
        for i in range(self.tbl.rowCount()):
            blob = " ".join(self.tbl.item(i, j).text() for j in range(self.tbl.columnCount())).lower().replace(",", "")
            self.tbl.setRowHidden(i, bool(t) and t not in blob)

    def approve_selected(self) -> int:
        if not self.res:
            return 0
        n = 0
        for c, p in zip(self.checks, self.pairs):
            if c.isChecked():
                manual_link(self.res, p["a"], p["b"], approve=True)
                n += 1
        if n:
            self.changed.emit(n)
        return n
