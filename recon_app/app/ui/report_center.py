"""مركز التقارير داخل التطبيق: معاينة كل ورقة كما ستظهر في Excel، واختيار ما يُصدَّر (ورقة أو أكثر أو الكل)."""
from __future__ import annotations

import csv
import os
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QFileDialog, QFrame, QHBoxLayout,
                               QLabel, QMessageBox, QPushButton, QScrollArea, QSizePolicy, QTableWidget,
                               QTableWidgetItem, QVBoxLayout, QWidget, QStyle, QStyledItemDelegate)

from ..exporter import ALL_SHEETS, SHEET_CATALOG, SHEET_NAME, export_report

MAX_PREVIEW_ROWS = 1200
PRESETS = [
    ("التقرير كاملاً", ALL_SHEETS),
    ("المتطابقات فقط", ["matched"]),
    ("للمراجعة والفروقات", ["review", "diff"]),
    ("الملخص والتحليل", ["summary", "analysis"]),
    ("ملفا المصدر", ["src_a", "src_b"]),
]
ICON = {"summary": "▦", "review": "⚑", "matched": "✔", "diff": "≠", "match": "⇄", "analysis": "◔",
        "settings": "⚙", "src_a": "▤", "src_b": "▤"}


class SheetRow(QFrame):
    """سطر ورقة: مربع اختيار للتصدير + نقر لمعاينتها."""
    clicked = Signal(str)
    toggled = Signal()

    def __init__(self, key: str, name: str, desc: str):
        super().__init__()
        self.key = key
        self.setObjectName("SheetRow")
        self.setCursor(Qt.PointingHandCursor)
        h = QHBoxLayout(self)
        h.setContentsMargins(12, 10, 12, 10)
        h.setSpacing(10)
        self.chk = QCheckBox()
        self.chk.setChecked(True)
        self.chk.setToolTip("ضمّ هذه الورقة إلى التصدير")
        self.chk.toggled.connect(lambda *_: self.toggled.emit())
        h.addWidget(self.chk)
        ic = QLabel(ICON.get(key, "▤"))
        ic.setFixedWidth(22)
        ic.setAlignment(Qt.AlignCenter)
        ic.setStyleSheet("color:#0D9488; font-size:16px; background:transparent;")
        h.addWidget(ic)
        v = QVBoxLayout()
        v.setSpacing(1)
        t = QLabel(name)
        t.setStyleSheet("font-weight:700; background:transparent;")
        d = QLabel(desc)
        d.setObjectName("Muted")
        d.setWordWrap(True)
        d.setStyleSheet("font-size:11px; background:transparent;")
        v.addWidget(t)
        v.addWidget(d)
        h.addLayout(v, 1)
        self.badge = QLabel("")
        self.badge.setObjectName("Badge")
        self.badge.setAlignment(Qt.AlignCenter)
        h.addWidget(self.badge, 0, Qt.AlignTop)
        self.set_active(False)

    def set_active(self, on: bool):
        self._on = on
        self.setStyleSheet(
            "#SheetRow { background: rgba(13,148,136,0.12); border: 1px solid #0D9488; border-radius: 12px; }"
            if on else
            "#SheetRow { background: transparent; border: 1px solid rgba(128,140,165,0.28); border-radius: 12px; }"
            "#SheetRow:hover { background: rgba(13,148,136,0.06); }")

    def mousePressEvent(self, e):
        self.clicked.emit(self.key)
        super().mousePressEvent(e)


def _fmt_cell(v, nf: str) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        if "%" in nf:
            return f"{v * 100:.1f}%"
        if "0.00" in nf:
            return f"{v:,.2f}"
        if "#,##0" in nf:
            return f"{v:,.0f}"
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    return str(v)


def _rgb(c) -> Optional[QColor]:
    try:
        rgb = c.rgb if c is not None else None
        if isinstance(rgb, str) and len(rgb) in (6, 8):
            return QColor("#" + rgb[-6:])
    except Exception:       # noqa: BLE001
        pass
    return None


class _FillDelegate(QStyledItemDelegate):
    """يرسم لون خلفية الخلية صراحةً (ورقة الأنماط العامة تتجاهل BackgroundRole)."""

    def paint(self, painter, option, index):
        b = index.data(Qt.BackgroundRole)
        if b is not None and not (option.state & QStyle.State_Selected):
            painter.fillRect(option.rect, b)
        super().paint(painter, option, index)


class ReportCenter(QWidget):
    """تبويب «التقرير»: قائمة الأوراق + معاينة + أزرار التصدير."""
    exportRequested = Signal(list)          # مفاتيح الأوراق المراد تصديرها

    def __init__(self):
        super().__init__()
        self.setObjectName("Page")
        self.ctx: Optional[Callable[[], dict]] = None
        self.current = "summary"
        self.dirty = True
        self._cache: dict[str, tuple] = {}
        self._rows: dict[str, SheetRow] = {}

        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 16, 0, 0)
        outer.setSpacing(18)

        # ---------------- العمود الأيمن: الأوراق والتصدير
        side = QFrame()
        side.setObjectName("Card")
        side.setFixedWidth(390)
        sv = QVBoxLayout(side)
        sv.setContentsMargins(18, 18, 18, 16)
        sv.setSpacing(10)
        t = QLabel("ماذا تريد أن تُصدّر؟")
        t.setObjectName("H2")
        sv.addWidget(t)
        s = QLabel("انقر على أي ورقة لمعاينتها هنا، وضع علامة ✔ على ما تريده في الملف. يمكنك تصدير ورقة واحدة فقط أو التقرير كاملاً.")
        s.setObjectName("Muted")
        s.setWordWrap(True)
        sv.addWidget(s)

        pr = QHBoxLayout()
        pr.setSpacing(6)
        pw = QWidget()
        pwl = QVBoxLayout(pw)
        pwl.setContentsMargins(0, 0, 0, 0)
        pwl.setSpacing(6)
        line = None
        for i, (name, keys) in enumerate(PRESETS):
            if i % 3 == 0:
                line = QHBoxLayout()
                line.setSpacing(6)
                pwl.addLayout(line)
            b = QPushButton(name)
            b.setObjectName("Chip")
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _=False, k=keys: self.select_only(k))
            line.addWidget(b)
        sv.addWidget(pw)

        sc = QScrollArea()
        sc.setWidgetResizable(True)
        sc.setFrameShape(QFrame.NoFrame)
        sc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        sc.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        host = QWidget()
        host.setObjectName("Plain")
        hv = QVBoxLayout(host)
        hv.setContentsMargins(0, 0, 4, 0)
        hv.setSpacing(8)
        for key, name, desc in SHEET_CATALOG:
            r = SheetRow(key, name, desc)
            r.clicked.connect(self.preview)
            r.toggled.connect(self._update_buttons)
            self._rows[key] = r
            hv.addWidget(r)
        hv.addStretch(1)
        sc.setWidget(host)
        sv.addWidget(sc, 1)

        self.btn_sel = QPushButton("")
        self.btn_sel.setObjectName("Primary")
        self.btn_sel.setCursor(Qt.PointingHandCursor)
        self.btn_sel.clicked.connect(lambda: self.exportRequested.emit(self.selected()))
        sv.addWidget(self.btn_sel)
        self.btn_all = QPushButton("تصدير التقرير كاملاً")
        self.btn_all.setObjectName("Ghost")
        self.btn_all.setCursor(Qt.PointingHandCursor)
        self.btn_all.clicked.connect(lambda: self.exportRequested.emit(self._available()))
        sv.addWidget(self.btn_all)
        outer.addWidget(side)

        # ---------------- المعاينة
        pv = QFrame()
        pv.setObjectName("Card")
        pl = QVBoxLayout(pv)
        pl.setContentsMargins(18, 16, 18, 14)
        pl.setSpacing(8)
        top = QHBoxLayout()
        tv = QVBoxLayout()
        tv.setSpacing(0)
        self.p_title = QLabel("")
        self.p_title.setObjectName("H2")
        self.p_sub = QLabel("")
        self.p_sub.setObjectName("Muted")
        tv.addWidget(self.p_title)
        tv.addWidget(self.p_sub)
        top.addLayout(tv, 1)
        self.b_one = QPushButton("تصدير هذه الورقة")
        self.b_one.setObjectName("Ghost")
        self.b_one.setCursor(Qt.PointingHandCursor)
        self.b_one.clicked.connect(lambda: self.exportRequested.emit([self.current]))
        self.b_csv = QPushButton("CSV")
        self.b_csv.setObjectName("Ghost")
        self.b_csv.setToolTip("تصدير هذه الورقة كملف CSV (يُفتح في أي برنامج)")
        self.b_csv.setCursor(Qt.PointingHandCursor)
        self.b_csv.clicked.connect(self.export_csv)
        top.addWidget(self.b_one)
        top.addWidget(self.b_csv)
        pl.addLayout(top)
        self.note = QLabel("")
        self.note.setObjectName("Muted")
        self.note.setWordWrap(True)
        self.note.hide()
        pl.addWidget(self.note)
        self.tbl = QTableWidget()
        self.tbl.setLayoutDirection(Qt.RightToLeft)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.tbl.setShowGrid(False)
        # المعاينة «ورقة Excel» بيضاء دائماً، فتبقى واضحة في الوضعين الفاتح والداكن
        self.tbl.setStyleSheet(
            "QTableWidget { background:#FFFFFF; alternate-background-color:#FFFFFF; color:#1F2937; border:1px solid #D5DBE5; }"
            "QTableWidget::item { padding:2px 6px; border:none; }"
            "QTableWidget::item:selected { background:#CCEFEA; }"
            "QTableWidget::item:hover { background:transparent; }")
        self.tbl.setItemDelegate(_FillDelegate(self.tbl))
        self.tbl.horizontalHeader().hide()
        self.tbl.verticalHeader().hide()
        self.tbl.setWordWrap(True)
        self.tbl.setTextElideMode(Qt.ElideNone)
        self.tbl.setFrameShape(QFrame.NoFrame)
        pl.addWidget(self.tbl, 1)
        outer.addWidget(pv, 1)

        self.set_available(set(ALL_SHEETS))
        self._update_buttons()

    # ------------------------------------------------------------------ سياق
    def set_context(self, fn: Callable[[], dict]):
        self.ctx = fn
        self.invalidate()

    def invalidate(self):
        self.dirty = True
        self._cache.clear()

    def ensure_fresh(self):
        if self.ctx is None or not self.dirty:
            return
        c = self.ctx()
        if not c or c.get("res") is None:
            return
        self.dirty = False
        avail = set(ALL_SHEETS)
        if c.get("table_a") is None:
            avail.discard("src_a")
        if c.get("table_b") is None:
            avail.discard("src_b")
        self.set_available(avail)
        self._fill_badges(c)
        self.preview(self.current if self.current in avail else "summary")

    def set_available(self, avail: set):
        self._avail = avail
        for k, r in self._rows.items():
            r.setVisible(k in avail)
            if k not in avail:
                r.chk.setChecked(False)
        self._update_buttons()

    def _available(self) -> list:
        return [k for k in ALL_SHEETS if k in self._avail]

    def _fill_badges(self, c: dict):
        from ..engine import STATUS_MATCHED
        res = c["res"]
        rows = res.rows
        n_ok = sum(1 for r in rows if r.status == STATUS_MATCHED)
        n_rev = sum(1 for r in rows if r.status != STATUS_MATCHED or r.priority == "منخفضة")
        n_diff = sum(1 for r in rows for _c, _k, amt in r.contribs if amt != 0)
        cnt = {"review": n_rev, "matched": n_ok, "diff": n_diff, "match": len(rows),
               "src_a": len(c["table_a"].rows) if c.get("table_a") else 0,
               "src_b": len(c["table_b"].rows) if c.get("table_b") else 0}
        for k, r in self._rows.items():
            n = cnt.get(k)
            r.badge.setText(f"{n:,}" if n is not None else "—")
            r.badge.setVisible(n is not None)

    # ------------------------------------------------------------------ اختيار
    def selected(self) -> list:
        return [k for k in ALL_SHEETS if k in self._rows and self._rows[k].isVisible() and self._rows[k].chk.isChecked()]

    def select_only(self, keys: list):
        for k, r in self._rows.items():
            r.chk.blockSignals(True)
            r.chk.setChecked(k in keys and k in self._avail)
            r.chk.blockSignals(False)
        self._update_buttons()

    def _update_buttons(self, *_):
        n = len(self.selected())
        self.btn_sel.setEnabled(n > 0)
        self.btn_sel.setText(f"  تصدير المحدد ({n} {'ورقة' if n == 1 else 'أوراق'}) إلى Excel" if n else "اختر ورقة للتصدير")

    # ------------------------------------------------------------------ معاينة
    def preview(self, key: str):
        self.current = key
        for k, r in self._rows.items():
            r.set_active(k == key)
        name = SHEET_NAME.get(key, key)
        desc = next((d for k, _n, d in SHEET_CATALOG if k == key), "")
        self.p_title.setText(name)
        self.p_sub.setText(desc)
        if self.ctx is None:
            return
        c = self.ctx()
        if not c or c.get("res") is None:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            data = self._cache.get(key)
            if data is None:
                data = self._build(key, c)
                self._cache[key] = data
            self._render(key, *data)
        finally:
            QApplication.restoreOverrideCursor()

    def _build(self, key: str, c: dict):
        """يولّد الورقة وحدها إلى ملف مؤقت ويقرأها (فتطابق المعاينة ما سيُصدَّر تماماً)."""
        import openpyxl
        fd, tmp = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        try:
            export_report(tmp, c["res"], c.get("table_a"), c.get("table_b"), c.get("skipped"),
                          c.get("names", ("", "")), sheets=[key])
            wb = openpyxl.load_workbook(tmp, data_only=True)
            ws = wb.active
            total = ws.max_row
            cells = []
            for row in ws.iter_rows(min_row=1, max_row=min(total, MAX_PREVIEW_ROWS)):
                line = []
                for cell in row:
                    fill = cell.fill.fgColor if cell.fill and cell.fill.fill_type == "solid" else None
                    line.append((_fmt_cell(cell.value, cell.number_format or ""), _rgb(fill),
                                 _rgb(cell.font.color) if cell.font and cell.font.color else None,
                                 bool(cell.font and cell.font.bold), (cell.font.size if cell.font else None),
                                 cell.alignment.horizontal if cell.alignment else None,
                                 bool(cell.hyperlink)))
                cells.append(line)
            widths = {}
            for col, dim in ws.column_dimensions.items():
                widths[dim.min or 0] = dim.width
                if dim.min and dim.max and dim.max > dim.min:
                    for k in range(dim.min, dim.max + 1):
                        widths[k] = dim.width
            heights = {r: d.height for r, d in ws.row_dimensions.items() if d.height}
            merges = [(m.min_row - 1, m.min_col - 1, m.max_row - m.min_row + 1, m.max_col - m.min_col + 1)
                      for m in ws.merged_cells.ranges if m.min_row <= MAX_PREVIEW_ROWS]
            charts = len(getattr(ws, "_charts", []))
            wb.close()
            return cells, widths, heights, merges, total, ws.max_column
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass

    def _render(self, key, cells, widths, heights, merges, total, ncols):
        t = self.tbl
        t.setUpdatesEnabled(False)
        t.clear()
        t.setRowCount(len(cells))
        t.setColumnCount(ncols)
        t.clearSpans()
        for j in range(ncols):
            w = widths.get(j + 1) or 9
            t.setColumnWidth(j, int(w * 7.2) + 8)
        for i, line in enumerate(cells):
            h = heights.get(i + 1)
            t.setRowHeight(i, min(64, max(22, int(h * 1.2))) if h else 24)
            for j, (txt, bg, fg, bold, size, ha, link) in enumerate(line):
                if not txt and bg is None:
                    continue
                it = QTableWidgetItem(txt)
                if bg is not None:
                    it.setBackground(QBrush(bg))
                if link:
                    it.setForeground(QBrush(QColor("#0F766E")))
                elif fg is not None:
                    it.setForeground(QBrush(fg))
                if bold or size:
                    f = QFont(it.font())
                    f.setBold(bool(bold))
                    if size and size > 11:
                        f.setPointSizeF(float(size) * 0.8)
                    it.setFont(f)
                it.setTextAlignment((Qt.AlignHCenter if ha == "center" else Qt.AlignLeading) | Qt.AlignVCenter)
                t.setItem(i, j, it)
        spanned = {(r0, c0) for r0, c0, _a, _b in merges}
        for i, line in enumerate(cells):                 # محاكاة امتداد النص الطويل فوق الخلايا الفارغة كما في Excel
            for j, (txt, bg, *_r) in enumerate(line):
                if not txt or (i, j) in spanned or len(txt) <= 12:
                    continue
                k = j + 1
                while k < ncols and not line[k][0] and line[k][1] is None:
                    k += 1
                if k - j > 1:
                    t.setSpan(i, j, 1, k - j)
        for r0, c0, rs, cs in merges:
            if r0 < len(cells) and (rs > 1 or cs > 1):
                t.setSpan(r0, c0, rs, cs)
        t.setUpdatesEnabled(True)
        t.scrollToTop()
        t.horizontalScrollBar().setValue(t.horizontalScrollBar().minimum())
        msgs = []
        if total > MAX_PREVIEW_ROWS:
            msgs.append(f"المعاينة تعرض أول {MAX_PREVIEW_ROWS:,} صف من {total:,} — التصدير يشمل الكل.")
        if key in ("summary", "analysis"):
            msgs.append("الرسوم البيانية تظهر في ملف Excel المُصدَّر.")
        self.note.setText("  ".join(msgs))
        self.note.setVisible(bool(msgs))

    # ------------------------------------------------------------------ CSV
    def export_csv(self):
        if self.ctx is None:
            return
        c = self.ctx()
        if not c or c.get("res") is None:
            return
        import openpyxl
        name = SHEET_NAME.get(self.current, self.current)
        path, _ = QFileDialog.getSaveFileName(
            self, "حفظ CSV", str(Path(c.get("last_dir") or Path.home()) / f"{name}_{datetime.now():%Y%m%d_%H%M}.csv"),
            "CSV (*.csv)")
        if not path:
            return
        fd, tmp = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        try:
            export_report(tmp, c["res"], c.get("table_a"), c.get("table_b"), c.get("skipped"),
                          c.get("names", ("", "")), sheets=[self.current])
            ws = openpyxl.load_workbook(tmp, data_only=True).active
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                for row in ws.iter_rows(values_only=True):
                    if any(v not in (None, "") for v in row):
                        w.writerow(["" if v is None else v for v in row])
        except PermissionError:
            QMessageBox.warning(self, "تعذّر الحفظ", "الملف مفتوح في برنامج آخر. أغلقه ثم أعد المحاولة.")
            return
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass
        w = self.window()
        if hasattr(w, "statusBar"):
            w.statusBar().showMessage(f"تم حفظ الملف: {path}", 8000)
