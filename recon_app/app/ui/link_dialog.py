"""نافذة «ربط بعملية مقابلة»: مرشّحات مرتّبة بالأنسب + بحث + إدخال رقم العملية مباشرة."""
from __future__ import annotations

from decimal import Decimal

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QAbstractItemView, QDialog, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
                               QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout)

from ..engine import STATUS_MISSING_B
from ..normalize import q2


def _fm(d) -> str:
    return "—" if d is None else f"{q2(Decimal(d)):,.2f}"


class LinkDialog(QDialog):
    def __init__(self, row, candidates, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ربط بعملية مقابلة")
        self.setLayoutDirection(Qt.RightToLeft)
        self.resize(980, 640)
        self.row = row
        self.cands = candidates            # [(MatchRow, score, why)]
        self.chosen = None
        self.approve = True
        mine_is_a = row.status == STATUS_MISSING_B
        mine = row.a if mine_is_a else row.b
        other_name = "ملف المحاسب" if mine_is_a else "ملف الصرافة"
        my_name = "الصرافة" if mine_is_a else "المحاسب"

        v = QVBoxLayout(self)
        v.setContentsMargins(22, 20, 22, 18)
        v.setSpacing(12)
        t = QLabel("ربط عمليتين هما في الحقيقة عملية واحدة")
        t.setObjectName("H2")
        v.addWidget(t)

        card = QLabel(
            f"<b>العملية الحالية (رقم {row.id}) — {'مفقودة في الصرافة' if mine_is_a else 'مفقودة في ملف المحاسب'}</b><br>"
            f"المبلغ: <b>{_fm(mine.amount)} {mine.currency}</b> &nbsp;·&nbsp; "
            f"التاريخ: {mine.date or '—'} &nbsp;·&nbsp; الصف {mine.row}<br>"
            f"<span style='color:gray'>{' '.join((mine.desc or '').split())[:170]}</span>")
        card.setObjectName("Warn")
        card.setWordWrap(True)
        card.setTextFormat(Qt.RichText)
        v.addWidget(card)

        sub = QLabel(f"اختر العملية المقابلة من «{other_name}» (مرتّبة بالأنسب أولاً)، أو اكتب رقمها مباشرة:")
        sub.setObjectName("Muted")
        sub.setWordWrap(True)
        v.addWidget(sub)

        bar = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("🔍  ابحث بالمبلغ أو الاسم أو أي كلمة من البيان…")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self._filter)
        bar.addWidget(self.search, 1)
        self.idbox = QLineEdit()
        self.idbox.setPlaceholderText("أو اكتب رقم العملية #")
        self.idbox.setMaximumWidth(190)
        self.idbox.returnPressed.connect(self._by_id)
        bar.addWidget(self.idbox)
        go = QPushButton("اذهب")
        go.setObjectName("Ghost")
        go.clicked.connect(self._by_id)
        bar.addWidget(go)
        v.addLayout(bar)

        self.tbl = QTableWidget(len(candidates), 6)
        self.tbl.setHorizontalHeaderLabels(["#", "الملاءمة", "التاريخ", "المبلغ", "البيان", "لماذا مقترحة؟"])
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setAlternatingRowColors(True)
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(4, QHeaderView.Stretch)
        for c, w in ((0, 60), (1, 80), (2, 100), (3, 130), (5, 220)):
            self.tbl.setColumnWidth(c, w)
        for i, (cand, score, why) in enumerate(candidates):
            rec = cand.b if mine_is_a else cand.a
            vals = [str(cand.id), f"{score}%", rec.date.strftime("%Y-%m-%d") if rec.date else "—",
                    f"{_fm(rec.amount)} {rec.currency}".strip(), " ".join((rec.desc or "").split())[:140],
                    "، ".join(why)]
            for j, val in enumerate(vals):
                it = QTableWidgetItem(val)
                it.setTextAlignment(Qt.AlignCenter if j in (0, 1, 2, 3) else Qt.AlignVCenter | Qt.AlignRight)
                if j == 1:
                    it.setForeground(QColor("#1F9D55" if score >= 75 else ("#D99100" if score >= 45 else "#8A93A3")))
                    f = it.font()
                    f.setBold(True)
                    it.setFont(f)
                self.tbl.setItem(i, j, it)
            self.tbl.setRowHeight(i, 34)
        self.tbl.itemDoubleClicked.connect(lambda *_: self._accept(True))
        v.addWidget(self.tbl, 1)

        self.info = QLabel("")
        self.info.setObjectName("Muted")
        self.info.setWordWrap(True)
        v.addWidget(self.info)

        foot = QHBoxLayout()
        self.b_ok = QPushButton("✔  ربط واعتماد (تنتقل إلى المتطابقة)")
        self.b_ok.setObjectName("Good")
        self.b_only = QPushButton("ربط فقط (للمراجعة لاحقاً)")
        self.b_only.setObjectName("Ghost")
        cancel = QPushButton("إلغاء")
        cancel.setObjectName("Ghost")
        for b in (self.b_ok, self.b_only, cancel):
            b.setCursor(Qt.PointingHandCursor)
        self.b_ok.setEnabled(False)
        self.b_only.setEnabled(False)
        self.b_ok.clicked.connect(lambda: self._accept(True))
        self.b_only.clicked.connect(lambda: self._accept(False))
        cancel.clicked.connect(self.reject)
        foot.addWidget(self.b_ok)
        foot.addWidget(self.b_only)
        foot.addStretch(1)
        foot.addWidget(cancel)
        v.addLayout(foot)

        self.tbl.itemSelectionChanged.connect(self._sel)
        if candidates and candidates[0][1] >= 45:
            self._select(0)
        else:
            self.info.setText("لا توجد مرشّحة قوية — جرّب البحث أو اكتب رقم العملية.")

    # ------------------------------------------------------------------
    def _filter(self, text: str):
        t = text.strip().lower()
        for i in range(self.tbl.rowCount()):
            blob = " ".join((self.tbl.item(i, j).text() if self.tbl.item(i, j) else "") for j in range(6)).lower()
            blob = blob.replace(",", "")
            self.tbl.setRowHidden(i, bool(t) and t.replace(",", "") not in blob)

    def _by_id(self):
        t = self.idbox.text().strip().lstrip("#")
        if not t.isdigit():
            self.info.setText("اكتب رقم العملية كما يظهر في عمود # (أرقام فقط).")
            return
        for i, (cand, _s, _w) in enumerate(self.cands):
            if cand.id == int(t):
                self.search.clear()
                self.tbl.setRowHidden(i, False)
                self._select(i)
                self.tbl.scrollToItem(self.tbl.item(i, 0))
                return
        self.info.setText(f"العملية رقم {t} ليست من العمليات التي يمكن ربطها بهذه العملية "
                          "(يجب أن تكون من الجهة الأخرى وفي حالة «مفقودة»).")

    def _select(self, i: int):
        self.tbl.clearSelection()
        self.tbl.setCurrentCell(i, 0)
        self.tbl.selectRow(i) if not self.tbl.selectionModel().selectedRows() else None

    def _sel(self):
        rows = self.tbl.selectionModel().selectedRows()
        ok = bool(rows)
        self.b_ok.setEnabled(ok)
        self.b_only.setEnabled(ok)
        if ok:
            cand, score, why = self.cands[rows[0].row()]
            self.chosen = cand
            self.info.setText(f"المحدّدة: العملية رقم {cand.id}  ·  ملاءمة {score}%"
                              + (f"  ·  {'، '.join(why)}" if why else ""))

    def _accept(self, approve: bool):
        if self.chosen is None:
            return
        self.approve = approve
        self.accept()
