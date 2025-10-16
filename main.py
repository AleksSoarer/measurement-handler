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
from odf.style import Style, TableCellProperties
from odf.table import Table, TableRow, TableCell
from odf.text import P

# ---- Colors ----
GREEN = QColor("#C6EFCE")   # soft green
RED   = QColor("#FFC7CE")   # soft red
BLUE  = QColor("#9DC3E6")   # kept for backward compat
WHITE = QColor("#FFFFFF")

# ---- Layout sizes ----
HDR_PANEL_HEIGHT = 60
TOL_PANEL_HEIGHT = 56
INFO_COL_WIDTH   = 220  # ширина левого фиксированного столбца

# ---- Special rows (0-based in MAIN table) ----
HEADER_ROWS    = [1, 2]  # «шапка», которую показываем сверху в отдельном виджете
NOMINAL_ROW    = 4       # 5-я строка: номинал (информативно)
TOL_ROW        = 5       # 6-я строка: допуск (редактируется в панели)
FIRST_DATA_ROW = 6       # данные с 7-й строки

MAX_CELLS = 200_000


# ---- Helpers ----
def has_letters(s: str) -> bool:
    return bool(re.search(r'[A-Za-zА-Яа-я]', s or ""))

def has_digits(s: str) -> bool:
    return any(ch.isdigit() for ch in (s or ""))

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


class MiniOdsEditor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Мини-редактор ODS (фикс. шапка + допуски + фикс. 1-й столбец)")
        self.resize(1250, 820)

        root = QVBoxLayout(self)

        # --- Controls ---
        ctrl = QHBoxLayout()
        ctrl.addWidget(QLabel("Колонки (X):"))
        self.sb_cols = QSpinBox(); self.sb_cols.setRange(1, 2000); self.sb_cols.setValue(8)
        ctrl.addWidget(self.sb_cols)

        ctrl.addWidget(QLabel("Строки (Y):"))
        self.sb_rows = QSpinBox(); self.sb_rows.setRange(1, 5000); self.sb_rows.setValue(12)
        ctrl.addWidget(self.sb_rows)

        self.btn_build = QPushButton("Создать таблицу"); self.btn_build.clicked.connect(self.build_table)
        ctrl.addWidget(self.btn_build)

        self.btn_open = QPushButton("Открыть .ods"); self.btn_open.clicked.connect(self.open_ods)
        ctrl.addWidget(self.btn_open)

        self.btn_save = QPushButton("Сохранить в .ods"); self.btn_save.clicked.connect(self.save_to_ods)
        ctrl.addWidget(self.btn_save)

        ctrl.addStretch()
        root.addLayout(ctrl)

        # ======= TOP PANELS (stacked vertically) =======
        # HEADER panel (rows 1–2)
        self.header_table = QTableWidget(len(HEADER_ROWS), 0, self)
        self._setup_top_table(self.header_table, HDR_PANEL_HEIGHT, font_inc=1.5)
        self.header_table.cellChanged.connect(self.on_hdr_cell_changed)

        # TOLERANCE panel (row 5)
        self.tolerance_table = QTableWidget(1, 0, self)
        self._setup_top_table(self.tolerance_table, TOL_PANEL_HEIGHT, font_inc=2.0)
        self.tolerance_table.cellChanged.connect(self.on_tol_cell_changed)

        # ======= CENTER AREA: left fixed column + right main stack =======
        center = QHBoxLayout()
        root.addLayout(center, stretch=1)

        # LEFT fixed column (row info)
        self.info_table = QTableWidget(0, 1, self)
        self._setup_info_table()

        # RIGHT side: vertical stack of header + tolerance + main
        right_stack = QVBoxLayout()

        # add widgets
        right_stack.addWidget(self.header_table)
        right_stack.addWidget(self.tolerance_table)

        # MAIN table
        self.table = QTableWidget(0, 0, self)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.verticalHeader().setVisible(False)
        self.table.cellChanged.connect(self.on_cell_changed)
        right_stack.addWidget(self.table)

        # put into center
        center.addWidget(self.info_table)
        # thin separator
        sep = QFrame(); sep.setFrameShape(QFrame.VLine); sep.setFrameShadow(QFrame.Sunken)
        center.addWidget(sep)
        container = QWidget(); container.setLayout(right_stack)
        center.addWidget(container, stretch=1)

        # --- Sync scrollbars & sizes ---
        # Horizontal: sync header/tolerance with main
        self.table.horizontalScrollBar().valueChanged.connect(self.header_table.horizontalScrollBar().setValue)
        self.table.horizontalScrollBar().valueChanged.connect(self.tolerance_table.horizontalScrollBar().setValue)
        self.header_table.horizontalScrollBar().valueChanged.connect(self.table.horizontalScrollBar().setValue)
        self.tolerance_table.horizontalScrollBar().valueChanged.connect(self.table.horizontalScrollBar().setValue)
        # Vertical: sync info column with main
        self.table.verticalScrollBar().valueChanged.connect(self.info_table.verticalScrollBar().setValue)
        self.info_table.verticalScrollBar().valueChanged.connect(self.table.verticalScrollBar().setValue)
        # Column width sync (panels)
        self.table.horizontalHeader().sectionResized.connect(self._on_main_section_resized)
        # Row height sync (left info <-> main)
        self.table.verticalHeader().sectionResized.connect(self._on_main_row_height_changed)

        # Init
        self.build_table()

    # ---------- UI setup helpers ----------
    def _setup_top_table(self, tw: QTableWidget, height: int, font_inc: float = 0.0):
        tw.verticalHeader().setVisible(False)
        tw.horizontalHeader().setVisible(False)
        tw.setFixedHeight(height)
        # set reasonable row heights
        rows = max(1, tw.rowCount())
        for r in range(rows):
            tw.setRowHeight(r, max(28, height // rows - 2))
        # font & padding
        f = tw.font()
        f.setPointSizeF(f.pointSizeF() + font_inc)
        tw.setFont(f)
        tw.setStyleSheet("QTableWidget::item { padding: 6px; }")
        # edits & scroll
        tw.setEditTriggers(QAbstractItemView.AllEditTriggers)
        tw.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tw.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)

    def _setup_info_table(self):
        self.info_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.info_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.info_table.verticalHeader().setVisible(False)
        self.info_table.horizontalHeader().setVisible(False)
        self.info_table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.info_table.setColumnWidth(0, INFO_COL_WIDTH)
        self.info_table.cellChanged.connect(self.on_info_cell_changed)
        # белый фон всегда
        self.info_table.setStyleSheet("QTableWidget::item { background: white; padding: 4px; }")

    # ---------- Coloring rules ----------
    def _get_tol(self, col):
        if col == 0:
            return None
        tol_it = self.table.item(TOL_ROW, col)
        tol = try_parse_float(tol_it.text()) if tol_it else None
        return tol

    def recolor_cell(self, it: QTableWidgetItem, row=None, col=None):
        if not it:
            return
        if row is None or col is None:
            idx = self.table.indexFromItem(it)
            row = idx.row(); col = idx.column()

        text_raw = it.text() or ""
        text = text_raw.strip()
        up = text.upper()

        # колонки 0 в MAIN скрыта и служит источником для info_table — держим белым
        if col == 0:
            it.setBackground(WHITE); return

        # строки шапки/номинал/допуск — не красим
        if row in HEADER_ROWS or row in (NOMINAL_ROW, TOL_ROW):
            it.setBackground(WHITE); return

        if up in ("N", "NM"):
            it.setBackground(RED); return
        if up == "Y":
            it.setBackground(GREEN); return

        f = try_parse_float(text)

        if (row >= FIRST_DATA_ROW) and (col > 0):
            tol = self._get_tol(col)
            if tol is not None and f is not None:
                it.setBackground(GREEN if abs(f) <= tol else RED)
                return

        # fallback
        if has_letters(text):
            it.setBackground(RED)
        elif has_digits(text):
            it.setBackground(GREEN)
        else:
            it.setBackground(WHITE)

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

    # ---------- Top panels sync ----------
    def _ensure_hdr_panel_cols(self):
        cols = self.table.columnCount()
        if self.header_table.columnCount() != cols:
            self.header_table.blockSignals(True)
            self.header_table.setColumnCount(cols)
            for r in range(self.header_table.rowCount()):
                for c in range(cols):
                    it = self.header_table.item(r, c)
                    if it is None:
                        self.header_table.setItem(r, c, QTableWidgetItem(""))
                    self.header_table.item(r, c).setTextAlignment(Qt.AlignCenter)
            self.header_table.blockSignals(False)
        # ширины (скрытый 0-й тоже синхронизируется, но мы его спрячем)
        for c in range(cols):
            self.header_table.setColumnWidth(c, self.table.columnWidth(c))
        # прячем 0-й столбец у панелей (его отображает info_table)
        self.header_table.setColumnHidden(0, True)

    def _sync_hdr_from_main(self):
        self._ensure_hdr_panel_cols()
        try:
            self.header_table.blockSignals(True)
            cols = self.table.columnCount()
            for i, r_main in enumerate(HEADER_ROWS):
                for c in range(cols):
                    src = self.table.item(r_main, c)
                    txt = src.text() if src else ""
                    it = self.header_table.item(i, c)
                    if it is None:
                        it = QTableWidgetItem("")
                        self.header_table.setItem(i, c, it)
                    it.setTextAlignment(Qt.AlignCenter)
                    it.setText(txt)
                    it.setBackground(WHITE)
        finally:
            self.header_table.blockSignals(False)

    def on_hdr_cell_changed(self, row, col):
        if col < 0 or row < 0:
            return
        main_row = HEADER_ROWS[row]
        txt = self.header_table.item(row, col).text() if self.header_table.item(row, col) else ""
        it = self.table.item(main_row, col)
        if it is None:
            it = QTableWidgetItem(""); self.table.setItem(main_row, col, it)
        try:
            self.table.blockSignals(True)
            it.setText(txt); it.setBackground(WHITE)
        finally:
            self.table.blockSignals(False)
        # не перекрашиваем — эти строки не красятся

    def _ensure_tol_panel_cols(self):
        cols = self.table.columnCount()
        if self.tolerance_table.columnCount() != cols:
            self.tolerance_table.blockSignals(True)
            self.tolerance_table.setColumnCount(cols)
            for c in range(cols):
                it = self.tolerance_table.item(0, c)
                if it is None:
                    self.tolerance_table.setItem(0, c, QTableWidgetItem(""))
                self.tolerance_table.item(0, c).setTextAlignment(Qt.AlignCenter)
            self.tolerance_table.blockSignals(False)
        for c in range(cols):
            self.tolerance_table.setColumnWidth(c, self.table.columnWidth(c))
        self.tolerance_table.setColumnHidden(0, True)

    def _sync_tol_from_main(self):
        self._ensure_tol_panel_cols()
        try:
            self.tolerance_table.blockSignals(True)
            cols = self.table.columnCount()
            for c in range(cols):
                src = self.table.item(TOL_ROW, c)
                txt = src.text() if src else ""
                it = self.tolerance_table.item(0, c)
                if it is None:
                    it = QTableWidgetItem(""); self.tolerance_table.setItem(0, c, it)
                it.setTextAlignment(Qt.AlignCenter)
                it.setText(txt)
                it.setBackground(WHITE)
        finally:
            self.tolerance_table.blockSignals(False)

    def _on_main_section_resized(self, logicalIndex, oldSize, newSize):
        if logicalIndex < self.tolerance_table.columnCount():
            self.tolerance_table.setColumnWidth(logicalIndex, newSize)
        if logicalIndex < self.header_table.columnCount():
            self.header_table.setColumnWidth(logicalIndex, newSize)

    # ---------- Left info column sync ----------
    def _ensure_info_rows(self):
        rows = self.table.rowCount()
        if self.info_table.rowCount() != rows:
            self.info_table.blockSignals(True)
            self.info_table.setRowCount(rows)
            for r in range(rows):
                it = self.info_table.item(r, 0)
                if it is None:
                    self.info_table.setItem(r, 0, QTableWidgetItem(""))
                self.info_table.item(r, 0).setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                # выровняем высоту
                self.info_table.setRowHeight(r, self.table.rowHeight(r))
            self.info_table.blockSignals(False)

    def _sync_info_from_main(self):
        self._ensure_info_rows()
        try:
            self.info_table.blockSignals(True)
            rows = self.table.rowCount()
            for r in range(rows):
                src = self.table.item(r, 0)
                txt = src.text() if src else ""
                it = self.info_table.item(r, 0)
                if it is None:
                    it = QTableWidgetItem(""); self.info_table.setItem(r, 0, it)
                it.setText(txt)
                it.setBackground(WHITE)
        finally:
            self.info_table.blockSignals(False)

    def _on_main_row_height_changed(self, logicalIndex, oldSize, newSize):
        if 0 <= logicalIndex < self.info_table.rowCount():
            self.info_table.setRowHeight(logicalIndex, newSize)

    def on_info_cell_changed(self, row, col):
        # проброс текста в скрытую колонку 0 основной таблицы
        txt = self.info_table.item(row, col).text() if self.info_table.item(row, col) else ""
        it = self.table.item(row, 0)
        if it is None:
            it = QTableWidgetItem(""); self.table.setItem(row, 0, it)
        try:
            self.table.blockSignals(True)
            it.setText(txt); it.setBackground(WHITE)
        finally:
            self.table.blockSignals(False)
        # не нужно перекрашивать — это инфо-колонка

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

            # скрываем служебные строки в main (их показывают панели)
            for r in HEADER_ROWS:
                if r < rows:
                    self.table.setRowHidden(r, True)
            if rows > TOL_ROW:
                self.table.setRowHidden(TOL_ROW, True)

            # скрываем колонку 0 в main — её показывает info_table
            if cols > 0:
                self.table.setColumnHidden(0, True)

            # синхронизируем панели
            self._ensure_hdr_panel_cols(); self._sync_hdr_from_main()
            self._ensure_tol_panel_cols(); self._sync_tol_from_main()

            # левую колонку заполняем и выравниваем по высоте
            self._sync_info_from_main()

        finally:
            self.table.blockSignals(False)

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
        self.recheck_column(col)

    def on_cell_changed(self, row, col):
        # если изменили скрытые строки в main — синхронизируем панели
        if row in HEADER_ROWS:
            try:
                self.header_table.blockSignals(True)
                idx = HEADER_ROWS.index(row)
                src = self.table.item(row, col)
                txt = src.text() if src else ""
                it = self.header_table.item(idx, col)
                if it is None:
                    it = QTableWidgetItem(""); self.header_table.setItem(idx, col, it)
                it.setTextAlignment(Qt.AlignCenter); it.setText(txt); it.setBackground(WHITE)
            finally:
                self.header_table.blockSignals(False)
            return

        if row == TOL_ROW:
            try:
                self.tolerance_table.blockSignals(True)
                src = self.table.item(TOL_ROW, col)
                txt = src.text() if src else ""
                it = self.tolerance_table.item(0, col)
                if it is None:
                    it = QTableWidgetItem(""); self.tolerance_table.setItem(0, col, it)
                it.setTextAlignment(Qt.AlignCenter); it.setText(txt); it.setBackground(WHITE)
            finally:
                self.tolerance_table.blockSignals(False)
            self.recheck_column(col)
            return

        # если изменили скрытую колонку 0 main (через код) — обновим info_table
        if col == 0:
            try:
                self.info_table.blockSignals(True)
                src = self.table.item(row, 0)
                txt = src.text() if src else ""
                it = self.info_table.item(row, 0)
                if it is None:
                    it = QTableWidgetItem(""); self.info_table.setItem(row, 0, it)
                it.setText(txt)
            finally:
                self.info_table.blockSignals(False)
            return

        # обычные данные — раскрасить
        it = self.table.item(row, col)
        self.recolor_cell(it, row, col)

    # ---------- ODS I/O ----------
    def save_to_ods(self):
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить как…", "table.ods", "ODS (*.ods)")
        if not path: return

        doc = OpenDocumentSpreadsheet()
        style_green = Style(name="bgGreen", family="table-cell")
        style_green.addElement(TableCellProperties(backgroundcolor="#C6EFCE"))
        doc.automaticstyles.addElement(style_green)
        style_red = Style(name="bgRed", family="table-cell")
        style_red.addElement(TableCellProperties(backgroundcolor="#FFC7CE"))
        doc.automaticstyles.addElement(style_red)
        style_blue = Style(name="bgBlue", family="table-cell")
        style_blue.addElement(TableCellProperties(backgroundcolor="#9DC3E6"))
        doc.automaticstyles.addElement(style_blue)
        style_white = None  # default

        t = Table(name="Sheet1"); doc.spreadsheet.addElement(t)

        rows = self.table.rowCount(); cols = self.table.columnCount()
        for r in range(rows):
            tr = TableRow(); t.addElement(tr)
            for c in range(cols):
                it = self.table.item(r, c)
                text = it.text() if it else ""
                bg = it.background().color() if it else WHITE

                if bg == GREEN: stylename = style_green
                elif bg == RED: stylename = style_red
                elif bg == BLUE: stylename = style_blue
                else: stylename = style_white

                f = try_parse_float(text)
                if f is not None:
                    cell = TableCell(valuetype="float", value=f, stylename=stylename)
                    cell.addElement(P(text=str(f)))
                else:
                    cell = TableCell(valuetype="string", stylename=stylename)
                    cell.addElement(P(text=text))
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
                            it.setText(text)
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

        # скрыть служебные строки и колонку 0, синхронизировать панели
        if self.table.rowCount() > TOL_ROW:
            self.table.setRowHidden(TOL_ROW, True)
        for r in HEADER_ROWS:
            if r < self.table.rowCount():
                self.table.setRowHidden(r, True)
        if self.table.columnCount() > 0:
            self.table.setColumnHidden(0, True)

        self._ensure_hdr_panel_cols(); self._sync_hdr_from_main()
        self._ensure_tol_panel_cols(); self._sync_tol_from_main()
        self._sync_info_from_main()

        if truncated:
            QMessageBox.information(
                self,
                "Файл урезан",
                f"Загружено {use_rows}×{use_cols} из {content_rows}×{content_cols} "
                f"(лимит ≈ {MAX_CELLS:,} ячеек)."
            )


def main():
    app = QApplication(sys.argv)
    w = MiniOdsEditor()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
