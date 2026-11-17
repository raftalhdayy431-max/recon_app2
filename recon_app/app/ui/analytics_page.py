"""تبويب «التحليلات»: ملاحظات ذكية + توزيع الفروق + الأطراف الأكثر مشاكل + المحور الزمني."""
from __future__ import annotations

from PySide6.QtCore import QRectF, Qt, Signal
from PySide6.QtGui import QColor, QFont, QPainter
from PySide6.QtWidgets import (QFrame, QGridLayout, QHBoxLayout, QLabel, QPushButton, QScrollArea,
                               QSizePolicy, QVBoxLayout, QWidget)

from ..analytics import Analysis, _fmt, _fmt_cur
from .widgets import HBar, IconTile, KpiCard, add_shadow

LEVEL = {"high": ("#D64535", "مهم"), "mid": ("#E09F3E", "انتبه"), "info": ("#2563EB", "معلومة"),
         "good": ("#1F9D55", "جيد")}
PALETTE = ["#0D9488", "#2563EB", "#E09F3E", "#D64535", "#7C5CD6", "#1F9D55", "#64748B"]


def _card(title: str, sub: str = "") -> tuple[QFrame, QVBoxLayout]:
    f = QFrame()
    f.setObjectName("Card")
    v = QVBoxLayout(f)
    v.setContentsMargins(24, 20, 24, 20)
    v.setSpacing(10)
    t = QLabel(title)
    t.setObjectName("H2")
    v.addWidget(t)
    if sub:
        s = QLabel(sub)
        s.setObjectName("Muted")
        s.setWordWrap(True)
        v.addWidget(s)
    add_shadow(f, 22, 4, 22)
    f.layout().setAlignment(Qt.AlignTop)
    return f, v


def _clear(layout):
    while layout.count():
        it = layout.takeAt(0)
        w = it.widget()
        if w is not None:
            w.hide()
            w.setParent(None)
            w.deleteLater()
        elif it.layout() is not None:
            _clear(it.layout())


class Timeline(QWidget):
    """أعمدة: إجمالي العمليات (رمادي) والمشاكل (أحمر) لكل فترة."""

    def __init__(self):
        super().__init__()
        self.data: list = []
        self.unit = "يوم"
        self.setMinimumHeight(190)

    def set(self, data, unit):
        self.data, self.unit = data, unit
        self.update()

    def paintEvent(self, e):
        if not self.data:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        txt = self.palette().text().color()
        r = QRectF(self.rect()).adjusted(8, 8, -8, -26)
        n = len(self.data)
        mx = max(t for _, t, _ in self.data) or 1
        slot = r.width() / n
        bw = max(3.0, min(34.0, slot * 0.62))
        f = QFont(self.font())
        f.setPointSizeF(max(7.0, f.pointSizeF() - 1.5))
        p.setFont(f)
        step = max(1, int(n / max(1, r.width() // 46)))
        base = QColor(txt)
        base.setAlpha(36)
        for i, (lbl, tot, bad) in enumerate(self.data):
            x = r.left() + slot * i + (slot - bw) / 2
            h = r.height() * tot / mx
            p.setPen(Qt.NoPen)
            p.setBrush(base)
            p.drawRoundedRect(QRectF(x, r.bottom() - h, bw, h), 3, 3)
            if bad:
                hb = r.height() * bad / mx
                p.setBrush(QColor("#D64535"))
                p.drawRoundedRect(QRectF(x, r.bottom() - hb, bw, hb), 3, 3)
            if i % step == 0:
                p.setPen(txt)
                p.drawText(QRectF(x - 20, r.bottom() + 4, bw + 40, 18), Qt.AlignHCenter | Qt.AlignTop, lbl)


class AnalyticsView(QScrollArea):
    drill = Signal(dict)

    def __init__(self):
        super().__init__()
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        w = QWidget()
        w.setObjectName("Page")
        self.setWidget(w)
        self.v = QVBoxLayout(w)
        self.v.setContentsMargins(0, 18, 10, 12)
        self.v.setSpacing(20)

        self.kpi_row = QGridLayout()
        self.kpi_row.setSpacing(18)
        self.v.addLayout(self.kpi_row)

        self.c_find, self.l_find = _card("ملاحظات ذكية", "ما يستحق انتباهك أولاً، مرتّب بالأهم، مع زر يفتح العمليات المعنية مباشرة.")
        self.v.addWidget(self.c_find)

        row = QHBoxLayout()
        row.setSpacing(20)
        self.c_cat, self.l_cat = _card("من أين يأتي الفرق؟", "توزيع القيمة غير المطابقة حسب السبب.")
        self.c_par, self.l_par = _card("الأطراف الأكثر مشاكل", "الأسماء ذات أكثر العمليات غير المطابقة.")
        row.addWidget(self.c_cat, 1)
        row.addWidget(self.c_par, 1)
        self.v.addLayout(row)

        self.c_time, self.l_time = _card("متى تظهر المشاكل؟", " ")
        self.time_sub = self.c_time.findChildren(QLabel)[1]
        self.tl = Timeline()
        self.l_time.addWidget(self.tl)
        self.legend_t = QLabel("▮ إجمالي العمليات      <span style='color:#D64535'>▮</span> غير المطابقة")
        self.legend_t.setObjectName("Muted")
        self.legend_t.setTextFormat(Qt.RichText)
        self.l_time.addWidget(self.legend_t)
        self.v.addWidget(self.c_time)
        self.v.addStretch(1)

    # ------------------------------------------------------------------
    def _bar_row(self, label: str, value_txt: str, frac: float, color: str, sub: str = "", click=None):
        box = QWidget()
        g = QVBoxLayout(box)
        g.setContentsMargins(0, 2, 0, 2)
        g.setSpacing(3)
        top = QHBoxLayout()
        a = QLabel(label)
        a.setWordWrap(True)
        a.setStyleSheet("font-weight:600; background: transparent;")
        b = QLabel(value_txt)
        b.setStyleSheet("background: transparent;")
        top.addWidget(a, 1)
        top.addWidget(b)
        g.addLayout(top)
        bar = HBar()
        bar.set(max(2, int(frac * 100)), color)
        g.addWidget(bar)
        if sub:
            s = QLabel(sub)
            s.setObjectName("Muted")
            s.setWordWrap(True)
            g.addWidget(s)
        if click:
            box.setCursor(Qt.PointingHandCursor)
            box.mousePressEvent = lambda e, c=click: c()
            box.setToolTip("انقر لعرض العمليات")
        return box

    def show_analysis(self, an: Analysis):
        # ---- مؤشرات
        _clear(self.kpi_row)
        cards = []
        k = KpiCard("عمليات تحتاج متابعة", "#D64535", icon_kind="alert")
        k.set(f"{an.issues_total:,}", f"من أصل {an.rows_total:,} عملية")
        cards.append(k)
        for cur, e in sorted(an.exposure.items(), key=lambda kv: -kv[1]["abs"])[:3]:
            k = KpiCard(f"القيمة غير المطابقة ({cur or '—'})", "#0D9488", icon_kind="percent")
            net = e["net"]
            side = "محاسب أعلى" if net > 0 else ("صرافة أعلى" if net < 0 else "متوازن")
            k.set(_fmt(e["abs"]), f"{e['n']} عملية · الصافي {_fmt(abs(net))} ({side})")
            cards.append(k)
        for i, c in enumerate(cards):
            self.kpi_row.addWidget(c, 0, i)
            self.kpi_row.setColumnStretch(i, 1)

        # ---- ملاحظات
        for wdg in getattr(self, "_find_widgets", []):
            wdg.hide()
            wdg.setParent(None)
            wdg.deleteLater()
        self._find_widgets = []
        for f in an.findings:
            col, tag = LEVEL.get(f.level, LEVEL["info"])
            row = QFrame()
            row.setObjectName("CardFlat")
            h = QHBoxLayout(row)
            h.setContentsMargins(14, 12, 14, 12)
            h.setSpacing(12)
            dot = QLabel(tag)
            dot.setAlignment(Qt.AlignCenter)
            dot.setFixedWidth(58)
            dot.setStyleSheet(f"background:{col}; color:white; border-radius:10px; padding:3px 0; font-weight:700;")
            h.addWidget(dot, 0, Qt.AlignTop)
            tv = QVBoxLayout()
            tv.setSpacing(2)
            t = QLabel(f.title)
            t.setObjectName("H3")
            t.setWordWrap(True)
            d = QLabel(f.detail)
            d.setWordWrap(True)
            d.setObjectName("Muted")
            tv.addWidget(t)
            tv.addWidget(d)
            h.addLayout(tv, 1)
            if f.action:
                b = QPushButton(f.action_label)
                b.setObjectName("Ghost")
                b.setCursor(Qt.PointingHandCursor)
                b.clicked.connect(lambda _=False, a=f.action: self.drill.emit(a))
                h.addWidget(b, 0, Qt.AlignVCenter)
            self.l_find.addWidget(row)
            self._find_widgets.append(row)

        # ---- التصنيفات
        for wdg in getattr(self, "_cat_widgets", []):
            wdg.hide(); wdg.setParent(None); wdg.deleteLater()
        self._cat_widgets = []
        by_cur: dict = {}
        for c in an.categories:
            by_cur.setdefault(c.currency, []).append(c)
        main = max(by_cur, key=lambda cur: sum(c.amount for c in by_cur[cur])) if by_cur else None
        for cur, lst in sorted(by_cur.items(), key=lambda kv: -sum(c.amount for c in kv[1])):
            tot = sum(c.amount for c in lst) or 1
            if len(by_cur) > 1:
                hd = QLabel(f"عملة {cur or '—'}")
                hd.setObjectName("H3")
                self.l_cat.addWidget(hd)
                self._cat_widgets.append(hd)
            for i, c in enumerate(lst[:6]):
                frac = float(c.amount / tot)
                w = self._bar_row(c.category, f"{_fmt(c.amount)}  ({frac * 100:.0f}٪)", frac, PALETTE[i % len(PALETTE)],
                                  f"{c.count} عملية")
                self.l_cat.addWidget(w)
                self._cat_widgets.append(w)
        if not by_cur:
            e = QLabel("لا توجد فروق.")
            e.setObjectName("Muted")
            self.l_cat.addWidget(e)
            self._cat_widgets.append(e)

        # ---- الأطراف
        for wdg in getattr(self, "_par_widgets", []):
            wdg.hide(); wdg.setParent(None); wdg.deleteLater()
        self._par_widgets = []
        if not an.has_parties:
            e = QLabel("لا يوجد عمود اسم في الملفات، لذا لا يمكن تحليل الأطراف.")
            e.setObjectName("Muted")
            e.setWordWrap(True)
            self.l_par.addWidget(e)
            self._par_widgets.append(e)
        else:
            mx = max((p.issues for p in an.parties), default=1)
            for p in an.parties[:8]:
                w = self._bar_row(p.party, f"{p.issues} من {p.total}", p.issues / mx, "#D64535",
                                  _fmt_cur(p.exposure), click=lambda n=p.party: self.drill.emit({"search": n}))
                self.l_par.addWidget(w)
                self._par_widgets.append(w)
            if not an.parties:
                e = QLabel("كل الأطراف مطابقة.")
                e.setObjectName("Muted")
                self.l_par.addWidget(e)
                self._par_widgets.append(e)

        # ---- الزمن
        self.c_time.setVisible(an.has_dates and len(an.timeline) > 1)
        self.time_sub.setText(f"عدد العمليات ومنها غير المطابقة لكل {an.timeline_unit} (حسب تاريخ العملية).")
        self.tl.set(an.timeline, an.timeline_unit)
