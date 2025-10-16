import sys
import re
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QSpinBox, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog,
    QMessageBox, QAbstractItemView, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

# ODS
from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.style import Style, TableCellProperties, TextProperties
from odf.table import Table, TableRow, TableCell
from odf.text import P

import os, tempfile
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtGui import QPainter
from PyQt5.QtCore import QRect
from PyQt5.QtWidgets import QTableView, QListView, QScrollArea
from pypdf import PdfReader, PdfWriter

# ---- Colors ----
GREEN = QColor("#1A8830")   # ok data
RED   = QColor("#8D1926")   # bad data
BLUE  = QColor("#265C8F")   # good data
WHITE = QColor("#FFFFFF")   #
BLACK = QColor("#000000")   # NoMeasure
TEXT = QColor("#000000")    #

# ---- Layout sizes ----
HDR_PANEL_HEIGHT = 140
TOL_PANEL_HEIGHT = 70
INFO_COL_WIDTH   = 200  # ширина левого фиксированного столбца

# ---- Special rows (0-based in MAIN table) ----
HEADER_ROWS    = [3, 4]  # «шапка», показываем сверху в отдельном виджете
NOMINAL_ROW    = 4       # 5-я строка: номинал (информативно)
TOL_ROW        = 5       # 6-я строка: допуск (редактируется в панели)
FIRST_DATA_ROW = 6       # данные с 7-й строки

MAX_CELLS = 200_000

# ---- Export font size (для ODS и PDF) ----
EXPORT_FONT_PT = 24.0   # меняй одно число: шрифт в сохраняемых файлах

# ---- UI font size (только виджетам на экране) ----
UI_FONT_PT = 10.0


# ---- Helpers ----
"""
def has_letters(s: str) -> bool:
    # Любая буква: латиница, кириллица и прочие алфавиты
    return any(ch.isalpha() for ch in (s or ""))

def has_digits(s: str) -> bool:
    # Уже ок, но приведём к единому стилю
    return any(ch.isdigit() for ch in (s or ""))
"""
def _fmt_serial(s: str) -> str:
    """Вернуть '283' вместо '283.0' (и '283,0'). Остальное — без изменений."""
    f = try_parse_float(s)
    if f is not None:
        i = int(round(f))
        if abs(f - i) < 1e-9:
            return str(i)
    return s or ""

def try_parse_float(s: str):
    if s is None:
        return None
    s = s.strip()
    if not s:
        return None
    candidate = s.replace(',', '.')
    try:
        return float(candidate)
    except ValueError:
        return None

def _extract_text_from_cell(cell: TableCell) -> str:
    parts = []
    for p in cell.getElementsByType(P):
        for node in getattr(p, 'childNodes', []):
            data = getattr(node, 'data', None)
            if data:
                parts.append(str(data))
    text = "".join(parts).strip()
    if not text:
        v = cell.getAttribute('value')
        if v is not None:
            text = str(v)
    return text

def _cell_has_content(cell: TableCell) -> bool:
    v = cell.getAttribute('value')
    if v not in (None, ""):
        return True
    for p in cell.getElementsByType(P):
        for node in getattr(p, 'childNodes', []):
            if getattr(node, 'data', None):
                return True
    return False

def _sheet_content_bounds(sheet: Table):
    content_cols = 0
    content_rows = 0
    for row in sheet.getElementsByType(TableRow):
        rrep = int(row.getAttribute('numberrowsrepeated') or 1)
        row_has_content = False
        last_col_in_row = -1
        col_idx = 0
        for cell in row.getElementsByType(TableCell):
            crep = int(cell.getAttribute('numbercolumnsrepeated') or 1)
            if _cell_has_content(cell):
                row_has_content = True
                last_col_in_row = max(last_col_in_row, col_idx + crep - 1)
            col_idx += crep
        if row_has_content:
            content_rows += rrep
            content_cols = max(content_cols, last_col_in_row + 1)
    return content_rows, content_cols

def _sync_left_caption_height(self):
    # высота первой «рабочей» строки в правой таблице
    if self.table.rowCount() > FIRST_DATA_ROW:
        h = self.table.rowHeight(FIRST_DATA_ROW)
        self.info_main_caption.setFixedHeight(h)


class MiniOdsEditor(QWidget):
    def __init__(self):
        super().__init__()
        self._tol_cache = []
        self._in_cell_style = False
        self.setWindowTitle("Контроль допусков")
        self.resize(1280, 840)

        root = QVBoxLayout(self)

        # --- Controls ---
        ctrl = QHBoxLayout()
        
        #ctrl.addWidget(QLabel("Колонки (X):"))
        self.sb_cols = QSpinBox(); self.sb_cols.setRange(1, 2000); self.sb_cols.setValue(8)
        #ctrl.addWidget(self.sb_cols)

        #ctrl.addWidget(QLabel("Строки (Y):"))
        self.sb_rows = QSpinBox(); self.sb_rows.setRange(1, 5000); self.sb_rows.setValue(12)
        #ctrl.addWidget(self.sb_rows)

        #self.btn_build = QPushButton("Создать таблицу"); self.btn_build.clicked.connect(self.build_table)
        #ctrl.addWidget(self.btn_build)

        self.btn_open = QPushButton("Открыть .ods"); self.btn_open.clicked.connect(self.open_ods)
        ctrl.addWidget(self.btn_open)

        self.btn_save = QPushButton("Сохранить в .ods"); self.btn_save.clicked.connect(self.save_to_ods)
        ctrl.addWidget(self.btn_save)

        self.btn_export_merged = QPushButton("PDF: внешний + таблица")
        self.btn_export_merged.setToolTip("Склеить: сначала выбранный PDF, затем вся таблица одним листом")
        self.btn_export_merged.clicked.connect(self.export_pdf_with_prefix_tableonly)
        ctrl.addWidget(self.btn_export_merged)

        ctrl.addStretch()
        root.addLayout(ctrl)

        # ======= TOP PANELS =======
        # left fixed header info (rows 1–2 of col0)
        self.info_header_table = QTableWidget(len(HEADER_ROWS), 1, self)
        self._setup_left_fixed_table(self.info_header_table, HDR_PANEL_HEIGHT)

        # center header (rows 1–2, all columns)
        self.header_table = QTableWidget(len(HEADER_ROWS), 0, self)
        self._setup_top_table(self.header_table, HDR_PANEL_HEIGHT, font_inc=1.5)
        self.header_table.cellChanged.connect(self.on_hdr_cell_changed)

        # left fixed tolerance info (row 5 of col0)
        self.info_tol_table = QTableWidget(1, 1, self)
        self._setup_left_fixed_table(self.info_tol_table, TOL_PANEL_HEIGHT)

        # center tolerance (row 5, all columns)
        self.tolerance_table = QTableWidget(1, 0, self)
        self._setup_top_table(self.tolerance_table, TOL_PANEL_HEIGHT, font_inc=2.0)
        self.tolerance_table.cellChanged.connect(self.on_tol_cell_changed)

   


        # left main info (col0 for all rows, except hidden 1,2,5 — мы их тоже скрываем здесь)
        self.info_main_table = QTableWidget(0, 1, self)
        self._setup_left_main_table(self.info_main_table)

        # Заголовок для левого фикс-столбца (залипает) — СОЗДАЁМ ДО добавления в layout
        self.info_main_caption = QTableWidget(1, 1, self)
        self.info_main_caption.verticalHeader().setVisible(False)
        self.info_main_caption.horizontalHeader().setVisible(False)
        self.info_main_caption.setFixedHeight(32)
        self.info_main_caption.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.info_main_caption.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.info_main_caption.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.info_main_caption.setColumnWidth(0, INFO_COL_WIDTH)  # ширина сразу здесь
        cap_item = QTableWidgetItem("Серийный/измерение")
        cap_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.info_main_caption.setItem(0, 0, cap_item)
        self.info_main_caption.setFrameStyle(QFrame.NoFrame)
        self.info_main_caption.setShowGrid(False)

        self.info_main_caption.setStyleSheet(
            "QTableWidget::item { background: white; font-weight: 600; padding: 6px; }"
        )

        # RIGHT side: vertical stack of header + tolerance + main
        right_stack = QVBoxLayout()
        # Для левой части создаём такую же «стек»-колонку
        left_stack = QVBoxLayout()

        # нет промежутков и полей — полосы «прилипают»
        for lay in (left_stack, right_stack):
            lay.setSpacing(0)
            lay.setContentsMargins(0, 0, 0, 0)

                # подпись для нижней строки ("Не в допуске")
        self.info_oos_caption = QTableWidget(1, 1, self)
        self.info_oos_caption.verticalHeader().setVisible(False)
        self.info_oos_caption.horizontalHeader().setVisible(False)
        self.info_oos_caption.setFixedHeight(32)
        self.info_oos_caption.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.info_oos_caption.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.info_oos_caption.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.info_oos_caption.setColumnWidth(0, INFO_COL_WIDTH)
        oos_cap_item = QTableWidgetItem("Не в допуске")
        oos_cap_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.info_oos_caption.setItem(0, 0, oos_cap_item)
        self.info_oos_caption.setFrameStyle(QFrame.NoFrame)
        self.info_oos_caption.setShowGrid(False)
        self.info_oos_caption.setStyleSheet(
            "QTableWidget::item { background: white; font-weight: 600; padding: 6px; }"
        )

        # полоса счётчиков "Не в допуске" под основной таблицей
        self.oos_table = QTableWidget(1, 0, self)
        self._setup_top_table(self.oos_table, height=34, font_inc=0.0)
        self.oos_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.oos_table.setFrameStyle(QFrame.NoFrame)
        self.oos_table.setShowGrid(False)
        self.oos_table.setStyleSheet("QTableWidget::item { background: white; font-weight: 600; }")


        # ======= CENTER AREA: left fixed main info + right main =======
        left_stack = QVBoxLayout()
        right_stack = QVBoxLayout()
        for lay in (left_stack, right_stack):
            lay.setSpacing(0)
            lay.setContentsMargins(0, 0, 0, 0)

        # LEFT
        left_stack.addWidget(self.info_header_table)
        left_sep1 = QFrame(); left_sep1.setFrameShape(QFrame.HLine); left_sep1.setFrameShadow(QFrame.Sunken)
        left_stack.addWidget(left_sep1)
        left_stack.addWidget(self.info_tol_table)
        left_sep2 = QFrame(); left_sep2.setFrameShape(QFrame.HLine); left_sep2.setFrameShadow(QFrame.Sunken)
        left_stack.addWidget(left_sep2)
        left_stack.addWidget(self.info_main_caption)
        left_stack.addWidget(self.info_main_table)
        left_stack.addWidget(self.info_oos_caption)

        # RIGHT
        right_stack.addWidget(self.header_table)
        right_sep1 = QFrame(); right_sep1.setFrameShape(QFrame.HLine); right_sep1.setFrameShadow(QFrame.Sunken)
        right_stack.addWidget(right_sep1)
        right_stack.addWidget(self.tolerance_table)
        self.order_table = QTableWidget(1, 0, self)
        self._setup_top_table(self.order_table, height=34, font_inc=0.0)
        self.order_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.order_table.setFrameStyle(QFrame.NoFrame)
        self.order_table.setShowGrid(False)
        self.order_table.setStyleSheet("QTableWidget::item { background: white; font-weight: 600; }")
        right_stack.addWidget(self.order_table)

        self.table = QTableWidget(0, 0, self)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setVisible(False)
        self.table.cellChanged.connect(self.on_cell_changed)
        right_stack.addWidget(self.table, 1)

        right_stack.addWidget(self.oos_table)

        


        # Оборачиваем стэки в виджеты
        left_container = QWidget(); left_container.setLayout(left_stack)
        right_container = QWidget(); right_container.setLayout(right_stack)

        # ЕДИНЫЙ печатаемый контейнер (его и будем рендерить в PDF)
        self.report_panel = QWidget()
        rep_layout = QHBoxLayout(self.report_panel)
        rep_layout.setContentsMargins(12, 12, 12, 12)
        rep_layout.setSpacing(12)
        rep_layout.addWidget(left_container)
        sep_mid = QFrame(); sep_mid.setFrameShape(QFrame.VLine); sep_mid.setFrameShadow(QFrame.Sunken)
        rep_layout.addWidget(sep_mid)
        rep_layout.addWidget(right_container, 1)

        # Добавляем в корневой лэйаут
        root.addWidget(self.report_panel, 1)

        f = self.table.font(); f.setPointSizeF(UI_FONT_PT); self.table.setFont(f)

        # левые подписи
        f = self.info_main_caption.font(); f.setPointSizeF(UI_FONT_PT); self.info_main_caption.setFont(f)
        f = self.info_oos_caption.font();  f.setPointSizeF(UI_FONT_PT); self.info_oos_caption.setFont(f)

        # левые таблицы (если вдруг не через _setup_* создавались)
        f = self.info_main_table.font(); f.setPointSizeF(UI_FONT_PT); self.info_main_table.setFont(f)
        f = self.info_header_table.font(); f.setPointSizeF(UI_FONT_PT); self.info_header_table.setFont(f)
        f = self.info_tol_table.font();    f.setPointSizeF(UI_FONT_PT); self.info_tol_table.setFont(f)

        # ==== TOTAL DEFECTS STRIP (отдельное окошко) ====
        total_bar = QHBoxLayout()
        total_bar.setContentsMargins(0, 6, 0, 0)
        total_bar.setSpacing(8)

        lbl_total = QLabel("Итого брак:")
        lbl_total.setStyleSheet("font-weight:600;")
        self.total_defects_lbl = QLabel("0")
        self.total_defects_lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.total_defects_lbl.setStyleSheet("font-weight:700; padding:4px 8px; border:1px solid #ccc; border-radius:6px;")

        total_bar.addStretch(1)
        total_bar.addWidget(lbl_total)
        total_bar.addWidget(self.total_defects_lbl)
        root.addLayout(total_bar)

        # --- Sync scrollbars & sizes ---
        # Horizontal: sync center header/tolerance with main
        self.table.horizontalScrollBar().valueChanged.connect(self.header_table.horizontalScrollBar().setValue)
        self.table.horizontalScrollBar().valueChanged.connect(self.tolerance_table.horizontalScrollBar().setValue)
        #self.header_table.horizontalScrollBar().valueChanged.connect(self.table.horizontalScrollBar().setValue)
        #self.tolerance_table.horizontalScrollBar().valueChanged.connect(self.table.horizontalScrollBar().setValue)
        # Горизонтальный скролл: main <-> order_row
        self.table.horizontalScrollBar().valueChanged.connect(self.order_table.horizontalScrollBar().setValue)
        self.order_table.horizontalScrollBar().valueChanged.connect(self.table.horizontalScrollBar().setValue)

        # Vertical: sync left main info with main
        self.table.verticalScrollBar().valueChanged.connect(self.info_main_table.verticalScrollBar().setValue)
        self.info_main_table.verticalScrollBar().valueChanged.connect(self.table.verticalScrollBar().setValue)

        # Подстраиваем значения при смене диапазона (чтобы не «уплывал» offset)
        self.table.verticalScrollBar().rangeChanged.connect(
            lambda _min, _max: self.info_main_table.verticalScrollBar().setValue(
                self.table.verticalScrollBar().value()
            )
        )
        self.info_main_table.verticalScrollBar().rangeChanged.connect(
            lambda _min, _max: self.table.verticalScrollBar().setValue(
                self.info_main_table.verticalScrollBar().value()
            )
        )
        # Column width sync (center panels)
        self.table.horizontalHeader().sectionResized.connect(self._on_main_section_resized)
        # Row height sync (left main info <-> main)
        self.table.verticalHeader().sectionResized.connect(self._on_main_row_height_changed)

        # Edits in left fixed tables should update main col0
        self.info_header_table.cellChanged.connect(self.on_info_header_cell_changed)
        self.info_tol_table.cellChanged.connect(self.on_info_tol_cell_changed)
        self.info_main_table.cellChanged.connect(self.on_info_main_cell_changed)

        # Init
        self.build_table()

        # main <-> oos_table (горизонтально)
        self.table.horizontalScrollBar().valueChanged.connect(self.oos_table.horizontalScrollBar().setValue)
        # oos_table без собственных полос, но связь в обе стороны не помешает
        self.oos_table.horizontalScrollBar().valueChanged.connect(self.table.horizontalScrollBar().setValue)

    # ---------- UI setup helpers ----------
    def _apply_service_row_visibility(self):
        """Спрятать все служебные строки (до FIRST_DATA_ROW) из нижних таблиц."""
        rows = self.table.rowCount()
        # в правой main-таблице
        for r in range(min(rows, FIRST_DATA_ROW)):
            self.table.setRowHidden(r, True)
        # в левой info_main — чтобы выравнивание не «плыло»
        for r in range(min(rows, FIRST_DATA_ROW)):
            if r < self.info_main_table.rowCount():
                self.info_main_table.setRowHidden(r, True)

    def _setup_top_table(self, tw: QTableWidget, height: int, font_inc: float = 0.0):
        tw.verticalHeader().setVisible(False)
        tw.horizontalHeader().setVisible(False)
        tw.setFixedHeight(height)
        rows = max(1, tw.rowCount())
        for r in range(rows):
            tw.setRowHeight(r, max(28, height // rows - 2))
        f = tw.font(); f.setPointSizeF(UI_FONT_PT + font_inc); tw.setFont(f)
        tw.setStyleSheet("QTableWidget::item { padding: 6px; }")
        tw.setEditTriggers(QAbstractItemView.AllEditTriggers)
        tw.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        # вот эти две строки обеспечат отсутствие полос и колёсика
        tw.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tw.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        tw.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)

    def _setup_left_fixed_table(self, tw: QTableWidget, height: int):
        tw.setColumnCount(1)
        tw.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tw.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tw.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        tw.verticalHeader().setVisible(False)
        tw.horizontalHeader().setVisible(False)
        tw.setEditTriggers(QAbstractItemView.AllEditTriggers)
        tw.setFixedHeight(height)
        for r in range(tw.rowCount()):
            tw.setRowHeight(r, max(28, height // max(1, tw.rowCount()) - 2))
        tw.setColumnWidth(0, INFO_COL_WIDTH)
        tw.setStyleSheet("QTableWidget::item { background: white; padding: 4px; }")
        f = tw.font(); f.setPointSizeF(UI_FONT_PT); tw.setFont(f)

    def _setup_left_main_table(self, tw: QTableWidget):
        tw.setColumnCount(1)
        tw.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tw.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        tw.verticalHeader().setVisible(False)
        tw.horizontalHeader().setVisible(False)
        tw.setEditTriggers(QAbstractItemView.AllEditTriggers)
        tw.setColumnWidth(0, INFO_COL_WIDTH)
        #tw.setStyleSheet("QTableWidget::item { background: white; padding: 4px; }")
        tw.setStyleSheet("QTableWidget::item { padding: 4px; }")
        f = tw.font(); f.setPointSizeF(UI_FONT_PT); tw.setFont(f)

    # ---------- Coloring rules ----------
    def _get_tol(self, col):
        if col <= 0:
            return None
        if 0 <= col < len(self._tol_cache):
            return self._tol_cache[col]
        return None

    def recolor_cell(self, it: QTableWidgetItem, row=None, col=None):
        if not it:
            return
        if self._in_cell_style:
            return
        self._in_cell_style = True
        try:
            if row is None or col is None:
                idx = self.table.indexFromItem(it)
                row = idx.row(); col = idx.column()

            text = (it.text() or "").strip()
            up = text.upper()

            # кол.0
            if col == 0:
                it.setBackground(WHITE); it.setForeground(TEXT); return

            # служебные
            if row in HEADER_ROWS or row in (NOMINAL_ROW, TOL_ROW):
                it.setBackground(WHITE); it.setForeground(TEXT); return

            if up == "NM":
                it.setBackground(BLACK); it.setForeground(WHITE); return
            if up == "N":
                it.setBackground(RED); it.setForeground(TEXT); return
            if up == "Y":
                it.setBackground(GREEN); it.setForeground(TEXT); return

            # числа и допуски
            f = try_parse_float(text)
            if (row >= FIRST_DATA_ROW) and (col > 0):
                tol = self._get_tol(col)
                if tol is not None and f is not None:
                    it.setBackground(BLUE if abs(f) <= tol else RED)
                    it.setForeground(TEXT if it.background().color() != BLACK else WHITE)
                    return

            # fallback без вызовов it.text() и хелперов
            if any(ch.isalpha() for ch in text):
                it.setBackground(RED);   it.setForeground(TEXT)
            elif any(ch.isdigit() for ch in text):
                it.setBackground(GREEN); it.setForeground(TEXT)
            else:
                it.setBackground(WHITE); it.setForeground(TEXT)
        finally:
            self._in_cell_style = False

    def recolor_all(self):
        try:
            self.table.blockSignals(True)
            for r in range(self.table.rowCount()):
                for c in range(self.table.columnCount()):
                    it = self.table.item(r, c)
                    if it:
                        self.recolor_cell(it, r, c)
        finally:
            self.table.blockSignals(False)

    def recheck_column(self, col: int):
        if col <= 0:
            return
        try:
            self.table.blockSignals(True)
            for r in range(FIRST_DATA_ROW, self.table.rowCount()):
                it = self.table.item(r, col)
                if it is None: continue
                self.recolor_cell(it, r, col)
        finally:
            self.table.blockSignals(False)

    def _rebuild_tol_cache(self):
        """Считать допуски из TOL_ROW в массив, чтобы раскраска не трогала таблицу."""
        cols = self.table.columnCount()
        self._tol_cache = [None] * cols
        for c in range(cols):
            if c == 0:
                continue
            it = self.table.item(TOL_ROW, c)
            txt = (it.text() if it else "") or ""
            self._tol_cache[c] = try_parse_float(txt.strip())

    
    def _is_row_defective(self, r: int) -> bool:
        """
        Строка бракована, если:
        - есть 'N' в любой ячейке c >= 1, ИЛИ
        - есть число вне допуска, ИЛИ
        - вся строка измерений пуста (для c >= 1).
        'NM' не считаем браком.
        """
        cols = self.table.columnCount()
        if cols <= 1:
            return True  # нет измерений — считаем браком

        has_any_value = False
        tol_cache = {}

        for c in range(1, cols):
            it = self.table.item(r, c)
            txt = (it.text() if it else "").strip()
            up = txt.upper()

            if txt != "":
                has_any_value = True

            # 'N' => брак (NM не считаем)
            if up == "N":
                return True
            if up == "NM":
                continue  # специально не считаем браком

            # проверка по допуску (если есть число и задан допуск)
            f = try_parse_float(txt)
            if f is not None:
                if c not in tol_cache:
                    tol_cache[c] = self._get_tol(c)
                tol = tol_cache[c]
                if tol is not None and abs(f) > tol:
                    return True

        # если значений не было вообще — пустая строка => брак
        return not has_any_value
    

    def _print_widget_to_single_pdf(self, pdf_path, widget):
        """Печатает ЛЮБОЙ контейнер (со всеми дочерними) в один лист PDF.
        Перед печатью временно разворачивает вложенные таблицы/списки,
        чтобы в PDF попал весь контент без полос прокрутки.
        """
        if widget is None:
            raise RuntimeError("Передан пустой виджет для печати")

        # 1) временно «распрямим» содержимое
        _restore = self._expand_children_for_print(widget)

        try:
            widget.adjustSize()
            QApplication.processEvents()

            content_w = max(1, widget.width())
            content_h = max(1, widget.height())

            printer = QPrinter(QPrinter.HighResolution)
            printer.setResolution(300)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(pdf_path)
            printer.setPaperSize(QPrinter.A4)
            printer.setFullPage(True)
            printer.setOrientation(QPrinter.Landscape if content_w >= content_h else QPrinter.Portrait)

            painter = QPainter(printer)
            if not painter.isActive():
                raise RuntimeError("Не удалось активировать QPainter для печати PDF")

            # Целевая область страницы (в пикселях устройства)
            target = printer.pageRect(QPrinter.DevicePixel)

            # Масштаб с сохранением пропорций + центрирование
            sx = target.width() / float(content_w)
            sy = target.height() / float(content_h)
            scale = min(sx, sy)

            view_w = int(content_w * scale)
            view_h = int(content_h * scale)
            offset_x = int((target.width() - view_w) / 2)
            offset_y = int((target.height() - view_h) / 2)

            painter.save()
            painter.translate(offset_x, offset_y)
            painter.scale(scale, scale)
            widget.render(painter, flags=QWidget.DrawChildren)
            painter.restore()
            painter.end()
        finally:
            # 2) вернуть всё как было
            try:
                _restore()
            except Exception:
                pass




    
    def export_pdf_with_prefix(self):
        """
        Новый вариант: [входной PDF (все страницы)] + [ВЕСЬ self.table одной страницей].
        Никаких панелей/капшенов — только «плоская» таблица, как в ODS.
        """
        in_path, _ = QFileDialog.getOpenFileName(self, "Выбери внешний PDF", "", "PDF Files (*.pdf)")
        if not in_path:
            return

        out_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как...", "merged.pdf", "PDF Files (*.pdf)")
        if not out_path:
            return

        # 1) соберём плоскую копию всей таблицы (все строки/столбцы, с цветами)
        flat = self._build_full_table_snapshot()

        # 2) распечатаем ЕЁ на один лист во временный PDF
        fd, tmp_pdf = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)


        f = flat.font()
        f.setPointSizeF(EXPORT_FONT_PT)
        flat.setFont(f)
        flat.resizeColumnsToContents()
        flat.resizeRowsToContents()


        try:
            self._print_table_to_single_pdf(tmp_pdf, flat)

            # 3) склейка: сначала входной, потом наш лист
            writer = PdfWriter()

            r1 = PdfReader(in_path)
            if getattr(r1, "is_encrypted", False):
                try:
                    r1.decrypt("")
                except Exception:
                    QMessageBox.warning(self, "Ошибка", "Входной PDF зашифрован.")
                    return
            for p in r1.pages:
                writer.add_page(p)

            r2 = PdfReader(tmp_pdf)
            for p in r2.pages:
                writer.add_page(p)

            with open(out_path, "wb") as f:
                writer.write(f)

            QMessageBox.information(self, "Готово", f"PDF сохранён:\n{out_path}")

        except Exception as e:
            QMessageBox.critical(self, "Провал", f"Не удалось собрать PDF:\n{e}")
        finally:
            try:
                os.remove(tmp_pdf)
            except Exception:
                pass
            # не забываем удалить временный виджет
            flat.deleteLater()


    def _build_offscreen_table_for_pdf(self) -> QTableWidget:
        """Полная автономная копия ВСЕЙ таблицы (включая кол.0 и служебные строки),
        с экспортным кеглем и автоподбором размеров, не зависящая от UI."""
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        tw = QTableWidget(rows, cols)          # parent=None — вне UI
        tw.verticalHeader().setVisible(False)
        tw.horizontalHeader().setVisible(False)
        tw.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tw.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tw.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tw.setWordWrap(False)

        # экспортный шрифт
        f = tw.font()
        f.setPointSizeF(float(EXPORT_FONT_PT))
        tw.setFont(f)

        # контент + стили
        for r in range(rows):
            for c in range(cols):
                src = self.table.item(r, c)
                txt = src.text() if src else ""
                it = QTableWidgetItem(txt)
                if src:
                    it.setBackground(src.background())
                    it.setForeground(src.foreground())
                    it.setTextAlignment(src.textAlignment())
                else:
                    it.setTextAlignment(Qt.AlignCenter)
                tw.setItem(r, c, it)

        # автоподбор под новый шрифт
        tw.resizeColumnsToContents()
        tw.resizeRowsToContents()

        # у 0-й колонки держим минимальную ширину инфо-колонки
        w0 = max(INFO_COL_WIDTH, tw.sizeHintForColumn(0))
        tw.setColumnWidth(0, w0)

        # на случай, если sizeHint где-то 0
        for c in range(cols):
            if tw.columnWidth(c) <= 0:
                tw.setColumnWidth(c, max(1, tw.sizeHintForColumn(c)))
        for r in range(rows):
            if tw.rowHeight(r) <= 0:
                tw.setRowHeight(r, max(1, tw.sizeHintForRow(r)))

        QApplication.processEvents()
        return tw

    def _print_table_to_single_pdf(self, pdf_path, table: QTableWidget):
        """
        Печать QTableWidget на ОДИН лист с масштабированием по большей стороне.
        Печатаем ВСЕ содержимое (без полос прокрутки).
        """
        vh = table.verticalHeader()
        hh = table.horizontalHeader()
        fw = table.frameWidth() * 2

        content_w = int(fw + vh.width() + sum(table.columnWidth(c) for c in range(table.columnCount())))
        content_h = int(fw + hh.height() + sum(table.rowHeight(r) for r in range(table.rowCount())))
        if content_w <= 0 or content_h <= 0:
            raise RuntimeError("Таблица пуста — печатать нечего.")

        printer = QPrinter(QPrinter.HighResolution)
        printer.setResolution(300)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(pdf_path)
        printer.setPaperSize(QPrinter.A4)
        printer.setFullPage(True)
        printer.setOrientation(QPrinter.Landscape if content_w >= content_h else QPrinter.Portrait)

        painter = QPainter(printer)
        if not painter.isActive():
            raise RuntimeError("Не удалось активировать QPainter для печати PDF")

        target = printer.pageRect(QPrinter.DevicePixel)

        old_size = table.size()
        old_hpol = table.horizontalScrollBarPolicy()
        old_vpol = table.verticalScrollBarPolicy()

        try:
            table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            table.resize(content_w, content_h)
            QApplication.processEvents()

            sx = target.width() / float(content_w)
            sy = target.height() / float(content_h)
            scale = min(sx, sy)

            view_w = max(1, int(content_w * scale))
            view_h = max(1, int(content_h * scale))
            offset_x = int((target.width() - view_w) / 2)
            offset_y = int((target.height() - view_h) / 2)

            painter.setViewport(QRect(offset_x, offset_y, view_w, view_h))
            painter.setWindow(QRect(0, 0, int(content_w), int(content_h)))

            table.render(painter, flags=QWidget.DrawChildren)
        finally:
            painter.end()
            table.resize(old_size)
            table.setHorizontalScrollBarPolicy(old_hpol)
            table.setVerticalScrollBarPolicy(old_vpol)


   


    def _expand_children_for_print(self, root_widget):
        """Убираем скроллы и растягиваем виджеты, чтобы в PDF попал весь контент."""
        saved = []

        # Таблицы (QTableWidget/QTableView)
        for tv in root_widget.findChildren(QTableView):
            st = {
                "w": tv,
                "size": tv.size(),
                "hpol": tv.horizontalScrollBarPolicy(),
                "vpol": tv.verticalScrollBarPolicy(),
            }
            saved.append(st)
            try:
                tv.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                tv.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                tv.resizeColumnsToContents()
                tv.resizeRowsToContents()

                vh = tv.verticalHeader(); hh = tv.horizontalHeader()
                fw = tv.frameWidth() * 2
                total_w = fw + vh.width() + sum(tv.columnWidth(c) for c in range(tv.model().columnCount()))
                total_h = fw + hh.height() + sum(tv.rowHeight(r) for r in range(tv.model().rowCount()))
                tv.resize(max(1, total_w), max(1, total_h))
            except Exception:
                pass

        # Списки
        for lv in root_widget.findChildren(QListView):
            st = {
                "w": lv,
                "size": lv.size(),
                "hpol": lv.horizontalScrollBarPolicy(),
                "vpol": lv.verticalScrollBarPolicy(),
            }
            saved.append(st)
            try:
                lv.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                lv.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                m = lv.model()
                rows = m.rowCount() if m is not None else 0
                base_h = lv.sizeHintForRow(0) if rows > 0 else lv.fontMetrics().height() + 6
                total_h = lv.frameWidth()*2 + base_h*max(1, rows)
                total_w = max(lv.width(), lv.viewport().sizeHint().width() + lv.frameWidth()*2)
                lv.resize(max(1, total_w), max(1, total_h))
            except Exception:
                pass

        # Скролл-области
        for sa in root_widget.findChildren(QScrollArea):
            inner = sa.widget()
            st = {
                "w": sa,
                "size": sa.size(),
                "hpol": sa.horizontalScrollBarPolicy(),
                "vpol": sa.verticalScrollBarPolicy(),
                "inner_size": inner.size() if inner else None,
            }
            saved.append(st)
            try:
                sa.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                sa.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                if inner:
                    inner.adjustSize()
                    inner.resize(inner.sizeHint())
                    sa.resize(inner.width() + sa.frameWidth()*2,
                            inner.height() + sa.frameWidth()*2)
            except Exception:
                pass

        def _restore():
            for st in saved:
                w = st["w"]
                try:
                    if hasattr(w, "setHorizontalScrollBarPolicy"):
                        w.setHorizontalScrollBarPolicy(st.get("hpol", Qt.ScrollBarAsNeeded))
                    if hasattr(w, "setVerticalScrollBarPolicy"):
                        w.setVerticalScrollBarPolicy(st.get("vpol", Qt.ScrollBarAsNeeded))
                    w.resize(st["size"])
                    if isinstance(w, QScrollArea) and st.get("inner_size") and w.widget():
                        w.widget().resize(st["inner_size"])
                except Exception:
                    pass

        return _restore


    
    def _row_is_empty_measurements(self, r: int) -> bool:
        """True, если во всех ячейках c>=1 пусто (игнорируем служебные строки)."""
        cols = self.table.columnCount()
        if cols <= 1 or r < FIRST_DATA_ROW:
            return True
        for c in range(1, cols):
            it = self.table.item(r, c)
            if it and (it.text() or "").strip() != "":
                return False
        return True


    def _recompute_total_defects(self):
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        if rows <= FIRST_DATA_ROW:
            self.total_defects_lbl.setText("0")
            return

        total_bad = 0
        self._in_cell_style = True
        try:
            for r in range(FIRST_DATA_ROW, rows):
                is_bad = self._is_row_defective(r)
                if is_bad:
                    total_bad += 1

                it_left = self.info_main_table.item(r, 0)
                if it_left is not None:
                    it_left.setBackground(RED if is_bad else WHITE)
                    it_left.setForeground(WHITE if is_bad else TEXT)

                it0 = self.table.item(r, 0)
                if it0 is not None:
                    it0.setBackground(RED if is_bad else WHITE)
                    it0.setForeground(WHITE if is_bad else TEXT)

                # заливка пустой бракованной строки
                is_empty_line = self._row_is_empty_measurements(r)
                if is_bad and is_empty_line:
                    for c in range(1, cols):
                        itc = self.table.item(r, c) or QTableWidgetItem("")
                        if self.table.item(r, c) is None:
                            self.table.setItem(r, c, itc)
                        itc.setBackground(RED)
                        itc.setForeground(TEXT)
                else:
                    for c in range(1, cols):
                        itc = self.table.item(r, c)
                        if itc and (itc.text() or "").strip() == "":
                            itc.setBackground(WHITE)
                            itc.setForeground(TEXT)
        finally:
            self._in_cell_style = False

        self.total_defects_lbl.setText(str(total_bad))

    # ---------- Panels/Info sync helpers ----------
    
    def _ensure_panel_cols(self):
        cols = self.table.columnCount()

        # center header/tol panels
        for tw in (self.header_table, self.tolerance_table):
            if tw.columnCount() != cols:
                tw.blockSignals(True)
                tw.setColumnCount(cols)
                # init cells
                for r in range(tw.rowCount()):
                    for c in range(cols):
                        it = tw.item(r, c)
                        if it is None:
                            tw.setItem(r, c, QTableWidgetItem(""))
                        tw.item(r, c).setTextAlignment(Qt.AlignCenter)
                tw.blockSignals(False)
            # widths
            for c in range(cols):
                tw.setColumnWidth(c, self.table.columnWidth(c))
            # hide col0 – его отображают left-виджеты
            if cols > 0:
                tw.setColumnHidden(0, True)

        # left main info rows = точно как в main
        rows = self.table.rowCount()
        if self.info_main_table.rowCount() != rows:
            self.info_main_table.blockSignals(True)
            self.info_main_table.setRowCount(rows)
            for r in range(rows):
                it = self.info_main_table.item(r, 0)
                if it is None:
                    self.info_main_table.setItem(r, 0, QTableWidgetItem(""))
                self.info_main_table.item(r, 0).setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.info_main_table.blockSignals(False)
        # высоты строк и скрытие 1,2,5 для точного совпадения
        for r in range(rows):
            self.info_main_table.setRowHeight(r, self.table.rowHeight(r))
        #for r in HEADER_ROWS + [TOL_ROW]:
        #    if r < rows:
        #        self.info_main_table.setRowHidden(r, True)

        # высоты строк
        for r in range(rows):
            self.info_main_table.setRowHeight(r, self.table.rowHeight(r))
        # скрываем все служебные (0..FIRST_DATA_ROW-1)
        for r in range(min(rows, FIRST_DATA_ROW)):
            self.info_main_table.setRowHidden(r, True)
        
        # left header & tol fixed tables – ширина колонки уже задана, высоты держим = панелям
        # ничего дополнительно делать не нужно здесь
        self._sync_order_row()

        # --- синхронизировать нижнюю полосу (oos_table) с main
        cols = self.table.columnCount()
        if self.oos_table.columnCount() != cols:
            self.oos_table.blockSignals(True)
            self.oos_table.setColumnCount(cols)
            for c in range(cols):
                it = self.oos_table.item(0, c)
                if it is None:
                    it = QTableWidgetItem("")
                    self.oos_table.setItem(0, c, it)
                self.oos_table.item(0, c).setTextAlignment(Qt.AlignCenter)
            self.oos_table.blockSignals(False)

        # ширины колонок = как у main
        for c in range(cols):
            self.oos_table.setColumnWidth(c, self.table.columnWidth(c))

        # скрываем 0-й столбец (описательный)
        if cols > 0:
            self.oos_table.setColumnHidden(0, True)

    def _sync_header_from_main(self):
        self.header_table.blockSignals(True)
        cols = self.table.columnCount()
        for i, r_main in enumerate(HEADER_ROWS):
            for c in range(cols):
                src = self.table.item(r_main, c)
                txt = src.text() if src else ""
                it = self.header_table.item(i, c)
                if it is None:
                    it = QTableWidgetItem(""); self.header_table.setItem(i, c, it)
                it.setTextAlignment(Qt.AlignCenter)
                it.setText(txt); it.setBackground(WHITE)
        self.header_table.blockSignals(False)

        # left header info (col0)
        self.info_header_table.blockSignals(True)
        for i, r_main in enumerate(HEADER_ROWS):
            src = self.table.item(r_main, 0)
            txt = src.text() if src else ""
            it = self.info_header_table.item(i, 0)
            if it is None:
                it = QTableWidgetItem(""); self.info_header_table.setItem(i, 0, it)
            it.setText(txt)
        self.info_header_table.blockSignals(False)

    def _sync_order_row(self):
        """Обновить 1-based нумерацию колонок в полосе order_table и подогнать ширины."""
        cols = self.table.columnCount()

        # кол-во колонок и ячейки
        if self.order_table.columnCount() != cols:
            self.order_table.blockSignals(True)
            self.order_table.setColumnCount(cols)
            for c in range(cols):
                it = self.order_table.item(0, c)
                if it is None:
                    it = QTableWidgetItem("")
                    self.order_table.setItem(0, c, it)
                self.order_table.item(0, c).setTextAlignment(Qt.AlignCenter)
            self.order_table.blockSignals(False)

        # ширины как у основной таблицы
        for c in range(cols):
            self.order_table.setColumnWidth(c, self.table.columnWidth(c))

        # 0-й столбец – описательный: скрываем и НЕ нумеруем
        if cols > 0:
            self.order_table.setColumnHidden(0, True)

        # проставить метки 1..N-1 для c >= 1
        self.order_table.blockSignals(True)
        for c in range(cols):
            it = self.order_table.item(0, c)
            if it is None:
                it = QTableWidgetItem("")
                self.order_table.setItem(0, c, it)
            it.setBackground(WHITE)       # никогда не красим
            it.setText("" if c == 0 else str(c))  # тут и есть нумерация с 1
        self.order_table.blockSignals(False)

        # 
        self._sync_bars_and_captions_height()


    

    
    def on_hdr_cell_changed(self, row, col):
        """
        Редактирование ячеек верхней шапки (header_table).
        Прокидываем текст в скрытые строки основной таблицы (HEADER_ROWS)
        и не запускаем перекраску (эти строки всегда белые).
        """
        if self._in_cell_style:
            return
        if row < 0 or col < 0:
            return

        main_row = HEADER_ROWS[row]

        src_item = self.header_table.item(row, col)
        txt = src_item.text() if src_item else ""

        it = self.table.item(main_row, col)
        if it is None:
            it = QTableWidgetItem("")
            self.table.setItem(main_row, col, it)

        try:
            self.table.blockSignals(True)
            it.setText(txt)
            it.setBackground(WHITE)  # служебные строки не красим
        finally:
            self.table.blockSignals(False)

        # Если редактировали 0-й столбец шапки — обновим левую фикс-таблицу для шапки
        if col == 0 and row < self.info_header_table.rowCount():
            try:
                self.info_header_table.blockSignals(True)
                left_it = self.info_header_table.item(row, 0)
                if left_it is None:
                    left_it = QTableWidgetItem("")
                    self.info_header_table.setItem(row, 0, left_it)
                left_it.setText(txt)
            finally:
                self.info_header_table.blockSignals(False)

    def _sync_tol_from_main(self):
        self.tolerance_table.blockSignals(True)
        cols = self.table.columnCount()
        for c in range(cols):
            src = self.table.item(TOL_ROW, c)
            txt = src.text() if src else ""
            it = self.tolerance_table.item(0, c)
            if it is None:
                it = QTableWidgetItem(""); self.tolerance_table.setItem(0, c, it)
            it.setTextAlignment(Qt.AlignCenter)
            it.setText(txt); it.setBackground(WHITE)
        self.tolerance_table.blockSignals(False)

        # left tol info (col0)
        self.info_tol_table.blockSignals(True)
        src = self.table.item(TOL_ROW, 0)
        txt = src.text() if src else ""
        it = self.info_tol_table.item(0, 0)
        if it is None:
            it = QTableWidgetItem(""); self.info_tol_table.setItem(0, 0, it)
        it.setText(txt)
        self.info_tol_table.blockSignals(False)

    def _sync_info_main_from_main(self):
        self.info_main_table.blockSignals(True)
        rows = self.table.rowCount()
        for r in range(rows):
            src = self.table.item(r, 0)
            txt = src.text() if src else ""
            it = self.info_main_table.item(r, 0)
            if it is None:
                it = QTableWidgetItem("")
                self.info_main_table.setItem(r, 0, it)
            it.setText(_fmt_serial(txt))  # показываем целые
        self.info_main_table.blockSignals(False)

    def _on_main_section_resized(self, logicalIndex, oldSize, newSize):
        if logicalIndex < self.tolerance_table.columnCount():
            self.tolerance_table.setColumnWidth(logicalIndex, newSize)
        if logicalIndex < self.header_table.columnCount():
            self.header_table.setColumnWidth(logicalIndex, newSize)
        if logicalIndex < self.order_table.columnCount():
            self.order_table.setColumnWidth(logicalIndex, newSize)
        if logicalIndex < self.oos_table.columnCount():
            self.oos_table.setColumnWidth(logicalIndex, newSize)

    def _on_main_row_height_changed(self, logicalIndex, oldSize, newSize):
        if 0 <= logicalIndex < self.info_main_table.rowCount():
            self.info_main_table.setRowHeight(logicalIndex, newSize)
        self._sync_bars_and_captions_height()

    def on_info_header_cell_changed(self, row, col):
        main_row = HEADER_ROWS[row]
        txt = self.info_header_table.item(row, col).text() if self.info_header_table.item(row, col) else ""
        it = self.table.item(main_row, 0)
        if it is None:
            it = QTableWidgetItem(""); self.table.setItem(main_row, 0, it)
        try:
            self.table.blockSignals(True)
            it.setText(txt); it.setBackground(WHITE)
        finally:
            self.table.blockSignals(False)

    def on_info_tol_cell_changed(self, row, col):
        txt = self.info_tol_table.item(row, col).text() if self.info_tol_table.item(row, col) else ""
        it = self.table.item(TOL_ROW, 0)
        if it is None:
            it = QTableWidgetItem(""); self.table.setItem(TOL_ROW, 0, it)
        try:
            self.table.blockSignals(True)
            it.setText(txt); it.setBackground(WHITE)
        finally:
            self.table.blockSignals(False)

    def on_info_main_cell_changed(self, row, col):
        raw = self.info_main_table.item(row, col).text() if self.info_main_table.item(row, col) else ""
        txt = _fmt_serial(raw)
        self.info_main_table.item(row, 0).setText(txt)
        it = self.table.item(row, 0)
        # обновляем левый виджет (на случай, если пользователь ввёл 283.0)
        self.info_main_table.item(row, 0).setText(txt)
        if it is None:
            it = QTableWidgetItem(""); self.table.setItem(row, 0, it)
        try:
            self.table.blockSignals(True)
            it.setText(txt) 
            #it.setBackground(WHITE)
        finally:
            self.table.blockSignals(False)

    def _sync_order_and_caption_height(self):
        """Высота полосы нумерации и левой подписи = высоте первой рабочей строки."""
        # высота первой рабочей строки (после служебных)
        if self.table.rowCount() > FIRST_DATA_ROW:
            h = self.table.rowHeight(FIRST_DATA_ROW)
        else:
            h = max(28, self.order_table.rowHeight(0))

        # полоса нумерации
        self.order_table.setRowHeight(0, h)
        self.order_table.setFixedHeight(h)  # без рамок достаточно ровно h

        # левая подпись «Серийный номер»
        self.info_main_caption.setFixedHeight(h)

    def _sync_bars_and_captions_height(self):
        """Высота полос (order/oos) и левых подписей = высоте первой рабочей строки."""
        if self.table.rowCount() > FIRST_DATA_ROW:
            h = self.table.rowHeight(FIRST_DATA_ROW)
        else:
            h = 34  # запасной

        # верхняя полоса нумерации
        self.order_table.setRowHeight(0, h)
        self.order_table.setFixedHeight(h)
        # нижняя полоса счётчиков
        self.oos_table.setRowHeight(0, h)
        self.oos_table.setFixedHeight(h)

        # левые подписи
        self.info_main_caption.setFixedHeight(h)
        self.info_oos_caption.setFixedHeight(h)

    # ---------- UI callbacks ----------
    def build_table(self):
        cols = max(self.sb_cols.value(), 1)
        rows = max(self.sb_rows.value(), FIRST_DATA_ROW + 1)
        try:
            self.table.blockSignals(True)
            self.table.setColumnCount(cols)
            self.table.setRowCount(rows)
            self.table.clearContents()

            for r in range(rows):
                for c in range(cols):
                    it = self.table.item(r, c)
                    if it is None:
                        it = QTableWidgetItem(""); self.table.setItem(r, c, it)
                    it.setTextAlignment(Qt.AlignCenter)
                    it.setBackground(WHITE)

            # скрываем служебные строки в main (их показывают верхние панели)
            """
            for r in HEADER_ROWS:
                if r < rows:
                    self.table.setRowHidden(r, True)
            if rows > TOL_ROW:
                self.table.setRowHidden(TOL_ROW, True)


            """
            # скрываем колонку 0 в main — её показывают левые таблицы
            if cols > 0:
                self.table.setColumnHidden(0, True)
            self._apply_service_row_visibility()

            # выровнять панели/левые таблицы и залить данными
            self._ensure_panel_cols()
            self._sync_header_from_main()
            self._sync_tol_from_main()
            self._rebuild_tol_cache()
            self._sync_info_main_from_main()
            self._sync_order_row()
            self._recompute_oos_counts()
            self._sync_bars_and_captions_height()
            self.recolor_all()
            self._recompute_total_defects()

        finally:
            self.table.blockSignals(False)
            self.table.horizontalScrollBar().setValue(0)
            self.order_table.horizontalScrollBar().setValue(0)

    def on_tol_cell_changed(self, row, col):
        if col < 0: return
        txt = self.tolerance_table.item(0, col).text() if self.tolerance_table.item(0, col) else ""
        it = self.table.item(TOL_ROW, col)
        if it is None:
            it = QTableWidgetItem(""); self.table.setItem(TOL_ROW, col, it)
        try:
            self.table.blockSignals(True)
            it.setText(txt); it.setBackground(WHITE)
        finally:
            self.table.blockSignals(False)
        self._sync_tol_from_main()  # обновим левый tol и панели (на случай правок)
        self._rebuild_tol_cache()
        self.recheck_column(col)
        self._recompute_oos_counts()
        self._recompute_total_defects()

    def on_cell_changed(self, row, col):
        # отражаем возможные изменения скрытых строк в панели/левых таблицах
        if row in HEADER_ROWS:
            self._sync_header_from_main()
            return
        if row == TOL_ROW:
            self._sync_tol_from_main()
            self._rebuild_tol_cache()
            self.recheck_column(col)
            return
        if col == 0:
            # изменили скрытую кол.0 в main (через код/загрузку)
            self._sync_info_main_from_main()
            return

        # обычные данные — раскрасить
        it = self.table.item(row, col)
        self.recolor_cell(it, row, col)

        if col > 0 and row >= FIRST_DATA_ROW:
            self._recompute_oos_counts()
            self._recompute_total_defects()
        elif col == 0 and row >= FIRST_DATA_ROW:
            # меняли серийник — могло стать "пустая строка" по факту? да, но считаем по c>=1.
            self._recompute_total_defects()

    # не в допуске 
    def _recompute_oos_counts(self):
        """Посчитать количество ЧИСЕЛ вне допуска по каждому столбцу (c >= 1)."""
        cols = self.table.columnCount()
        rows = self.table.rowCount()
        if cols == 0 or rows == 0:
            return

        # гарантируем структуру/ширины нижней полосы
        self._ensure_panel_cols()

        self.oos_table.blockSignals(True)
        for c in range(cols):
            if c == 0:
                # описательный столбец — пусто
                cell = self.oos_table.item(0, c)
                if cell is None:
                    cell = QTableWidgetItem("")
                    self.oos_table.setItem(0, c, cell)
                cell.setText("")
                cell.setBackground(WHITE)
                continue

            tol = self._get_tol(c)
            cnt = 0
            if tol is not None:
                for r in range(FIRST_DATA_ROW, rows):
                    it = self.table.item(r, c)
                    if not it:
                        continue
                    f = try_parse_float(it.text())
                    if f is None:
                        continue  # считаем ТОЛЬКО числа
                    if abs(f) > tol:
                        cnt += 1

            cell = self.oos_table.item(0, c)
            if cell is None:
                cell = QTableWidgetItem("")
                self.oos_table.setItem(0, c, cell)
            cell.setText(str(cnt) if c > 0 else "")
            cell.setBackground(WHITE)
        self.oos_table.blockSignals(False)

    # ---------- ODS I/O ----------
    def save_to_ods(self):
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить как…", "table.ods", "ODS (*.ods)")
        if not path: return


        #GREEN = QColor("#1A8830")   # soft green
        #RED   = QColor("#8D1926")   # soft red
        #BLUE  = QColor("#265C8F")   # kept for backward compat (открытие старых .ods)
        #WHITE = QColor("#FFFFFF")
        
        doc = OpenDocumentSpreadsheet()

        # Общие текстовые настройки для ячеек (шрифт)
        def _txt_props():
            return TextProperties(fontsize=f"{EXPORT_FONT_PT}pt")

        style_green = Style(name="cellGreen", family="table-cell")
        style_green.addElement(TableCellProperties(backgroundcolor="#1A8830"))
        style_green.addElement(_txt_props())
        doc.automaticstyles.addElement(style_green)

        style_red = Style(name="cellRed", family="table-cell")
        style_red.addElement(TableCellProperties(backgroundcolor="#8D1926"))
        style_red.addElement(_txt_props())
        doc.automaticstyles.addElement(style_red)

        style_blue = Style(name="cellBlue", family="table-cell")
        style_blue.addElement(TableCellProperties(backgroundcolor="#265C8F"))
        style_blue.addElement(_txt_props())
        doc.automaticstyles.addElement(style_blue)

        style_white = Style(name="cellWhite", family="table-cell")
        style_white.addElement(TableCellProperties(backgroundcolor="#FFFFFF"))
        style_white.addElement(_txt_props())
        doc.automaticstyles.addElement(style_white)

        t = Table(name="Sheet1"); doc.spreadsheet.addElement(t)
        
        rows = self.table.rowCount(); cols = self.table.columnCount()
        for r in range(rows):
            tr = TableRow(); t.addElement(tr)
            for c in range(cols):
                it = self.table.item(r, c)
                text = it.text() if it else ""
                bg = it.background().color() if it else WHITE
                if bg == GREEN:   stylename = style_green
                elif bg == RED:   stylename = style_red
                elif bg == BLUE:  stylename = style_blue
                else:             stylename = style_white  # теперь не None — чтобы применился размер шрифта

                f = try_parse_float(text)
                if f is not None:
                    if c == 0 and abs(f - int(round(f))) < 1e-9:
                        ival = int(round(f))
                        cell = TableCell(valuetype="float", value=ival, stylename=stylename)
                        cell.addElement(P(text=str(ival)))
                    else:
                        cell = TableCell(valuetype="float", value=f, stylename=stylename)
                        cell.addElement(P(text=str(f)))
                else:
                    # для строки в первом столбце тоже подчистим внешний вид
                    txt = _fmt_serial(text) if c == 0 else text
                    cell = TableCell(valuetype="string", stylename=stylename)
                    cell.addElement(P(text=txt))
                tr.addElement(cell)

        doc.save(path)
        self.btn_save.setText("Сохранено ✓")

    def open_ods(self):
        path, _ = QFileDialog.getOpenFileName(self, "Открыть…", "", "ODS (*.ods)")
        if not path: return

        doc = load(path)
        tables = doc.spreadsheet.getElementsByType(Table)
        if not tables: return
        sheet = tables[0]

        
        content_rows, content_cols = _sheet_content_bounds(sheet)
        if content_rows == 0 or content_cols == 0:
            try:
                self.table.blockSignals(True); self.table.setUpdatesEnabled(False)
                self.table.clearContents(); self.table.setRowCount(1); self.table.setColumnCount(1)
                self.sb_rows.setValue(1); self.sb_cols.setValue(1)
                it = QTableWidgetItem(""); it.setTextAlignment(Qt.AlignCenter); it.setBackground(WHITE)
                self.table.setItem(0, 0, it)
            finally:
                self.table.setUpdatesEnabled(True); self.table.blockSignals(False)
            return

        est_cells = content_rows * content_cols
        truncated = est_cells > MAX_CELLS
        use_cols = content_cols
        use_rows = content_rows if not truncated else max(1, MAX_CELLS // max(1, use_cols))

        try:
            self.table.blockSignals(True); self.table.setUpdatesEnabled(False)
            self.table.clearContents()
            final_rows = max(use_rows, FIRST_DATA_ROW + 1)
            self.table.setRowCount(final_rows); self.table.setColumnCount(use_cols)
            self.sb_rows.setValue(final_rows); self.sb_cols.setValue(use_cols)

            row_idx = 0
            for row in sheet.getElementsByType(TableRow):
                if row_idx >= use_rows: break
                rrep = int(row.getAttribute('numberrowsrepeated') or 1)

                template = []
                col_idx = 0
                for cell in row.getElementsByType(TableCell):
                    crep = int(cell.getAttribute('numbercolumnsrepeated') or 1)
                    vis = min(crep, max(0, use_cols - col_idx))
                    if vis > 0:
                        text = _extract_text_from_cell(cell)
                        template.append((text, vis))
                    col_idx += crep
                    if col_idx >= use_cols: break

                for _ in range(rrep):
                    if row_idx >= use_rows: break
                    c = 0
                    for text, vis in template:
                        for _k in range(vis):
                            it = self.table.item(row_idx, c)
                            if it is None:
                                it = QTableWidgetItem(""); self.table.setItem(row_idx, c, it)
                            it.setTextAlignment(Qt.AlignCenter)
                            text_to_set = _fmt_serial(text) if c == 0 else text
                            it.setText(text_to_set)
                            self.recolor_cell(it, row_idx, c)
                            c += 1
                            if c >= use_cols: break
                        if c >= use_cols: break
                    row_idx += 1

            # заполнить оставшиеся (если расширили вниз)
            for r in range(use_rows, final_rows):
                for c in range(use_cols):
                    it = self.table.item(r, c)
                    if it is None:
                        it = QTableWidgetItem(""); self.table.setItem(r, c, it)
                    it.setTextAlignment(Qt.AlignCenter); it.setBackground(WHITE)

        finally:
            self.table.setUpdatesEnabled(True); self.table.blockSignals(False)

        # скрыть служебные строки и колонку 0, синхронизировать панели и левые виджеты
        """
        if self.table.rowCount() > TOL_ROW:
            self.table.setRowHidden(TOL_ROW, True)
        for r in HEADER_ROWS:
            if r < self.table.rowCount():
                self.table.setRowHidden(r, True)
        if self.table.columnCount() > 0:
            self.table.setColumnHidden(0, True)
        """
        self._apply_service_row_visibility()
        if self.table.columnCount() > 0:
            self.table.setColumnHidden(0, True)
        
        self._ensure_panel_cols()
        self._sync_header_from_main()
        self._sync_tol_from_main()
        self._rebuild_tol_cache()
        self._sync_info_main_from_main()
        self._sync_order_row()
        # чтобы нумерация начиналась с 1 на экране
        self.table.horizontalScrollBar().setValue(0)
        self.order_table.horizontalScrollBar().setValue(0)
        self._recompute_oos_counts()
        self._sync_bars_and_captions_height()
        self.recolor_all()
        self._recompute_total_defects() 

        if truncated:
            QMessageBox.information(
                self,
                "Файл урезан",
                f"Загружено {use_rows}×{use_cols} из {content_rows}×{content_cols} "
                f"(лимит ≈ {MAX_CELLS:,} ячеек)."
            )

    def export_pdf_with_prefix_tableonly(self):
        in_path, _ = QFileDialog.getOpenFileName(self, "Выбери внешний PDF (пойдёт в начало)", "", "PDF Files (*.pdf)")
        if not in_path:
            return
        out_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как...", "merged.pdf", "PDF Files (*.pdf)")
        if not out_path:
            return

        # 1) offscreen-копия всей таблицы
        tw = self._build_offscreen_table_for_pdf()

        # 2) печать на один лист во временный PDF
        fd, tmp_pdf = tempfile.mkstemp(suffix=".pdf"); os.close(fd)
        try:
            self._print_table_to_single_pdf(tmp_pdf, tw)

            # 3) склейка: сначала входной PDF, затем наш лист
            writer = PdfWriter()
            r1 = PdfReader(in_path)
            if getattr(r1, "is_encrypted", False):
                try:
                    r1.decrypt("")
                except Exception:
                    QMessageBox.warning(self, "Ошибка", "Входной PDF зашифрован.")
                    return
            for p in r1.pages:
                writer.add_page(p)

            r2 = PdfReader(tmp_pdf)
            for p in r2.pages:
                writer.add_page(p)

            with open(out_path, "wb") as f:
                writer.write(f)

            QMessageBox.information(self, "Готово", f"PDF сохранён:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Не удалось собрать PDF", str(e))
        finally:
            try: os.remove(tmp_pdf)
            except Exception: pass
            tw.deleteLater()


    def _print_whole_table_to_single_pdf(self, pdf_path, table, font_pt):
        """Печатает ИМЕННО всю основную таблицу (включая скрытые служебные строки и колонку 0)
        на один лист PDF с масштабированием, не затрагивая UI."""
        if table.rowCount() == 0 or table.columnCount() == 0:
            raise RuntimeError("Таблица пуста — печатать нечего.")

        # --- сохранить состояние, чтобы потом вернуть ---
        hidden_rows = [table.isRowHidden(r) for r in range(table.rowCount())]
        col0_hidden = table.isColumnHidden(0)
        old_font = table.font()
        old_hpol = table.horizontalScrollBarPolicy()
        old_vpol = table.verticalScrollBarPolicy()
        old_size = table.size()

        try:
            # Показать всё: колонку 0 и служебные строки
            table.setColumnHidden(0, False)
            for r in range(table.rowCount()):
                table.setRowHidden(r, False)

            # Временный шрифт только для печати
            f = old_font
            f.setPointSizeF(float(font_pt))
            table.setFont(f)

            # Подогнать размеры ячеек под шрифт
            table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            table.resizeColumnsToContents()
            table.resizeRowsToContents()
            table.clearSpans()  # на всякий

            # Полный размер содержимого
            vh = table.verticalHeader(); hh = table.horizontalHeader(); fw = table.frameWidth() * 2
            content_w = int(fw + vh.width() + sum(table.columnWidth(c) for c in range(table.columnCount())))
            content_h = int(fw + hh.height() + sum(table.rowHeight(r) for r in range(table.rowCount())))
            content_w = max(1, content_w)
            content_h = max(1, content_h)
            table.resize(content_w, content_h)
            QApplication.processEvents()

            # Подготовка принтера
            printer = QPrinter(QPrinter.HighResolution)
            printer.setResolution(300)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(pdf_path)
            printer.setPaperSize(QPrinter.A4)
            printer.setFullPage(True)
            printer.setOrientation(QPrinter.Landscape if content_w >= content_h else QPrinter.Portrait)

            painter = QPainter(printer)
            if not painter.isActive():
                raise RuntimeError("Не удалось активировать QPainter для печати PDF")

            # Прямоугольник страницы в пикселях устройства
            target = printer.pageRect(QPrinter.DevicePixel)

            # Масштаб с сохранением пропорций
            sx = target.width() / float(content_w)
            sy = target.height() / float(content_h)
            scale = min(sx, sy)

            view_w = max(1, int(content_w * scale))
            view_h = max(1, int(content_h * scale))
            offset_x = int((target.width() - view_w) / 2)
            offset_y = int((target.height() - view_h) / 2)

            # Вьюпорт/оконная система — всё целиком на страницу
            painter.setViewport(offset_x, offset_y, view_w, view_h)
            painter.setWindow(0, 0, content_w, content_h)
            table.render(painter, flags=QWidget.DrawChildren)
            painter.end()

        finally:
            # --- вернуть состояние ---
            table.setFont(old_font)
            table.setHorizontalScrollBarPolicy(old_hpol)
            table.setVerticalScrollBarPolicy(old_vpol)
            table.resize(old_size)
            table.setColumnHidden(0, col0_hidden)
            for r, was_hidden in enumerate(hidden_rows):
                table.setRowHidden(r, was_hidden)


def main():
    app = QApplication(sys.argv)
    w = MiniOdsEditor()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
