import sys
import re
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QSpinBox, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog,
    QMessageBox, QAbstractItemView
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

# ODS
from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.style import Style, TableCellProperties
from odf.table import Table, TableRow, TableCell
from odf.text import P

# ---- Colors and limits ----
GREEN = QColor("#C6EFCE")   # soft green
RED   = QColor("#FFC7CE")   # soft red
BLUE  = QColor("#9DC3E6")   # kept for backward compatibility on open
WHITE = QColor("#FFFFFF")

# Header rows (0-based indices in MAIN table)
NOMINAL_ROW = 4        # 5-я строка: номинал (информативно)
TOL_ROW     = 5        # 6-я строка: допуск (редактируется через закреплённую панель)
FIRST_DATA_ROW = 6     # данные с 7-й строки

# Ряды, которые выводим в отдельной «шапке» (и скрываем в основной таблице)
HEADER_ROWS = [1, 2, 3, 4]   # 2-я и 3-я строки (0-based)

# Высоты закреплённых панелей
HDR_PANEL_HEIGHT = 320
TOL_PANEL_HEIGHT = 80

# Hard cap for table size on load (to avoid OOM)
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
    """Collect visible text from <text:p> nodes or fallback to office:value."""
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
    """Detect non-empty cell for bounds detection (fast check)."""
    v = cell.getAttribute('value')
    if v not in (None, ""):
        return True
    for p in cell.getElementsByType(P):
        for node in getattr(p, 'childNodes', []):
            if getattr(node, 'data', None):
                return True
    return False


def _sheet_content_bounds(sheet: Table):
    """
    Determine content area (rows, cols) ignoring trailing empty rows/cols,
    respecting number-rows/columns-repeated attributes without expansion.
    Returns (content_rows, content_cols).
    """
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
        self.setWindowTitle("Мини-редактор ODS (фикс. шапка + допуски)")
        self.resize(1150, 800)

        root = QVBoxLayout(self)

        # Controls
        ctrl = QHBoxLayout()
        ctrl.addWidget(QLabel("Колонки (X):"))
        self.sb_cols = QSpinBox()
        self.sb_cols.setRange(1, 2000)
        self.sb_cols.setValue(8)
        ctrl.addWidget(self.sb_cols)

        ctrl.addWidget(QLabel("Строки (Y):"))
        self.sb_rows = QSpinBox()
        self.sb_rows.setRange(1, 5000)
        self.sb_rows.setValue(12)
        ctrl.addWidget(self.sb_rows)

        self.btn_build = QPushButton("Создать таблицу")
        self.btn_build.clicked.connect(self.build_table)
        ctrl.addWidget(self.btn_build)

        self.btn_open = QPushButton("Открыть .ods")
        self.btn_open.clicked.connect(self.open_ods)
        ctrl.addWidget(self.btn_open)

        self.btn_save = QPushButton("Сохранить в .ods")
        self.btn_save.clicked.connect(self.save_to_ods)
        ctrl.addWidget(self.btn_save)

        ctrl.addStretch()
        root.addLayout(ctrl)

        # === HEADER PANEL: фиксированная 2-строчная шапка (строки 1–2 основной таблицы) ===
        self.header_table = QTableWidget(len(HEADER_ROWS), 0, self)
        self.header_table.verticalHeader().setVisible(False)
        self.header_table.horizontalHeader().setVisible(False)
        self.header_table.setFixedHeight(HDR_PANEL_HEIGHT)
        # комфортная высота строк
        for r in range(len(HEADER_ROWS)):
            self.header_table.setRowHeight(r, HDR_PANEL_HEIGHT // len(HEADER_ROWS))
        # читаемый шрифт и паддинги
        hdr_font = self.header_table.font()
        hdr_font.setPointSizeF(hdr_font.pointSizeF() + 1.5)
        self.header_table.setFont(hdr_font)
        self.header_table.setStyleSheet("QTableWidget::item { padding: 6px; }")
        # редактируемо
        self.header_table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.header_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.header_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.header_table.cellChanged.connect(self.on_hdr_cell_changed)
        root.addWidget(self.header_table)

        # --- Tolerance panel (fixed top 1-row table) ---
        self.tolerance_table = QTableWidget(1, 0, self)
        self.tolerance_table.verticalHeader().setVisible(False)
        self.tolerance_table.horizontalHeader().setVisible(False)
        self.tolerance_table.setFixedHeight(TOL_PANEL_HEIGHT)
        self.tolerance_table.setRowHeight(0, TOL_PANEL_HEIGHT - 8)
        tol_font = self.tolerance_table.font()
        tol_font.setPointSizeF(tol_font.pointSizeF() + 2)
        self.tolerance_table.setFont(tol_font)
        self.tolerance_table.setStyleSheet("QTableWidget::item { padding: 6px; }")
        self.tolerance_table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.tolerance_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tolerance_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.tolerance_table.cellChanged.connect(self.on_tol_cell_changed)
        root.addWidget(self.tolerance_table)

        # Main table
        self.table = QTableWidget(0, 0, self)
        self.table.cellChanged.connect(self.on_cell_changed)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.verticalHeader().setVisible(False)
        root.addWidget(self.table)

        # sync scrollbars & column widths (три виджета синхронно)
        self.table.horizontalScrollBar().valueChanged.connect(
            self.tolerance_table.horizontalScrollBar().setValue
        )
        self.table.horizontalScrollBar().valueChanged.connect(
            self.header_table.horizontalScrollBar().setValue
        )
        self.tolerance_table.horizontalScrollBar().valueChanged.connect(
            self.table.horizontalScrollBar().setValue
        )
        self.header_table.horizontalScrollBar().valueChanged.connect(
            self.table.horizontalScrollBar().setValue
        )
        self.table.horizontalHeader().sectionResized.connect(self._on_main_section_resized)

        # Initial table
        self.build_table()

    # ---- Coloring rules ----
    def _get_tol(self, col):
        """Return tolerance for given column or None. Skip col=0."""
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
            row = idx.row()
            col = idx.column()

        text_raw = it.text() or ""
        text = text_raw.strip()
        up = text.upper()

        # первый столбец (описания/имена) всегда белый
        if col == 0:
            it.setBackground(WHITE)
            return

        # строки шапки и служебные строки (не красим): белые
        if row in HEADER_ROWS or row in (NOMINAL_ROW, TOL_ROW):
            it.setBackground(WHITE)
            return

        # Маркеры соответствия/несоответствия (применяются только к данным)
        if up == "N" or up == "NM":
            it.setBackground(RED)
            return
        if up == "Y":
            it.setBackground(GREEN)
            return

        f = try_parse_float(text)

        # Только для данных (с 7-й строки) и если допуск в колонке задан
        if (row >= FIRST_DATA_ROW) and (col > 0):
            tol = self._get_tol(col)
            if tol is not None and f is not None:
                if abs(f) <= tol:
                    it.setBackground(GREEN)   # в допуске
                else:
                    it.setBackground(RED)     # вне допуска
                return

        # fallback для прочего (выше шапки данных быть не должно)
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
        """Repaint entire column after tolerance change."""
        if col <= 0:
            return
        try:
            self.table.blockSignals(True)
            for r in range(FIRST_DATA_ROW, self.table.rowCount()):
                it = self.table.item(r, col)
                if it is None:
                    continue
                self.recolor_cell(it, r, col)
        finally:
            self.table.blockSignals(False)

    # ---- HEADER PANEL helpers ----
    def _ensure_hdr_panel_cols(self):
        cols = self.table.columnCount()
        if self.header_table.columnCount() != cols:
            self.header_table.blockSignals(True)
            self.header_table.setColumnCount(cols)
            for r in range(len(HEADER_ROWS)):
                for c in range(cols):
                    it = self.header_table.item(r, c)
                    if it is None:
                        self.header_table.setItem(r, c, QTableWidgetItem(""))
                    self.header_table.item(r, c).setTextAlignment(Qt.AlignCenter)
            self.header_table.blockSignals(False)
        # синхронизируем ширины
        for c in range(cols):
            w = self.table.columnWidth(c)
            if self.header_table.columnWidth(c) != w:
                self.header_table.setColumnWidth(c, w)

    def _sync_hdr_from_main(self):
        """Скопировать строки 1–2 из основной таблицы в панель-шапку."""
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
                    # всегда белые
                    it.setBackground(WHITE)
        finally:
            self.header_table.blockSignals(False)

    def on_hdr_cell_changed(self, row, col):
        """Редактирование шапки: дублируем в скрытые строки основной таблицы."""
        if col < 0 or row < 0:
            return
        main_row = HEADER_ROWS[row]
        txt = self.header_table.item(row, col).text() if self.header_table.item(row, col) else ""
        it = self.table.item(main_row, col)
        if it is None:
            it = QTableWidgetItem("")
            self.table.setItem(main_row, col, it)
        try:
            self.table.blockSignals(True)
            it.setText(txt)
            it.setBackground(WHITE)
        finally:
            self.table.blockSignals(False)
        # перекраска не нужна — эти строки не красим и они скрыты

    # ---- TOLERANCE PANEL helpers ----
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
        # синхронизируем ширины
        for c in range(cols):
            w = self.table.columnWidth(c)
            if self.tolerance_table.columnWidth(c) != w:
                self.tolerance_table.setColumnWidth(c, w)

    def _sync_tol_from_main(self):
        """Скопировать значения допусков из скрытой строки основной таблицы в панель."""
        self._ensure_tol_panel_cols()
        try:
            self.tolerance_table.blockSignals(True)
            cols = self.table.columnCount()
            for c in range(cols):
                src = self.table.item(TOL_ROW, c)
                txt = src.text() if src else ""
                it = self.tolerance_table.item(0, c)
                if it is None:
                    it = QTableWidgetItem("")
                    self.tolerance_table.setItem(0, c, it)
                it.setTextAlignment(Qt.AlignCenter)
                it.setText(txt)
                it.setBackground(WHITE)
        finally:
            self.tolerance_table.blockSignals(False)

    def _on_main_section_resized(self, logicalIndex, oldSize, newSize):
        # Подгоняем ширину столбца обоих панелей под основной
        if logicalIndex < self.tolerance_table.columnCount():
            self.tolerance_table.setColumnWidth(logicalIndex, newSize)
        if logicalIndex < self.header_table.columnCount():
            self.header_table.setColumnWidth(logicalIndex, newSize)

    # ---- UI callbacks ----
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
                        it = QTableWidgetItem("")
                        self.table.setItem(r, c, it)
                    it.setTextAlignment(Qt.AlignCenter)
                    it.setBackground(WHITE)

            # панели: создать столбцы и синхронизировать содержимое
            self._ensure_hdr_panel_cols()
            self._ensure_tol_panel_cols()

            # скопировать из основной таблицы
            self._sync_hdr_from_main()
            self._sync_tol_from_main()

            # спрятать служебные строки в основной таблице
            for r in HEADER_ROWS:
                if r < rows:
                    self.table.setRowHidden(r, True)
            if rows > TOL_ROW:
                self.table.setRowHidden(TOL_ROW, True)

        finally:
            self.table.blockSignals(False)

    def on_tol_cell_changed(self, row, col):
        """Пользователь отредактировал допуск в панели — обновим скрытую строку и перекрасим колонку."""
        if col < 0:
            return
        txt = self.tolerance_table.item(0, col).text() if self.tolerance_table.item(0, col) else ""
        it = self.table.item(TOL_ROW, col)
        if it is None:
            it = QTableWidgetItem("")
            self.table.setItem(TOL_ROW, col, it)
        try:
            self.table.blockSignals(True)
            it.setText(txt)
            it.setBackground(WHITE)  # допуск — информативный
        finally:
            self.table.blockSignals(False)
        self.recheck_column(col)

    def on_cell_changed(self, row, col):
        # синхронизация назад, если вдруг пользователь отредактировал скрытые строки в main (через код/загрузку)
        if row in HEADER_ROWS:
            # обновим header_panel одну ячейку
            try:
                self.header_table.blockSignals(True)
                idx = HEADER_ROWS.index(row)
                src = self.table.item(row, col)
                txt = src.text() if src else ""
                it = self.header_table.item(idx, col)
                if it is None:
                    it = QTableWidgetItem("")
                    self.header_table.setItem(idx, col, it)
                it.setTextAlignment(Qt.AlignCenter)
                it.setText(txt)
                it.setBackground(WHITE)
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
                    it = QTableWidgetItem("")
                    self.tolerance_table.setItem(0, col, it)
                it.setTextAlignment(Qt.AlignCenter)
                it.setText(txt)
                it.setBackground(WHITE)
            finally:
                self.tolerance_table.blockSignals(False)
            self.recheck_column(col)
            return

        # обычные ячейки
        it = self.table.item(row, col)
        self.recolor_cell(it, row, col)

    # ---- ODS I/O ----
    def save_to_ods(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как…", "table.ods", "ODS (*.ods)"
        )
        if not path:
            return

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

        t = Table(name="Sheet1")
        doc.spreadsheet.addElement(t)

        rows = self.table.rowCount()
        cols = self.table.columnCount()

        for r in range(rows):
            tr = TableRow()
            t.addElement(tr)
            for c in range(cols):
                it = self.table.item(r, c)
                text = it.text() if it else ""
                bg = it.background().color() if it else WHITE

                if bg == GREEN:
                    stylename = style_green
                elif bg == RED:
                    stylename = style_red
                elif bg == BLUE:
                    stylename = style_blue
                else:
                    stylename = style_white

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
        if not path:
            return

        doc = load(path)
        tables = doc.spreadsheet.getElementsByType(Table)
        if not tables:
            return
        sheet = tables[0]

        content_rows, content_cols = _sheet_content_bounds(sheet)
        if content_rows == 0 or content_cols == 0:
            try:
                self.table.blockSignals(True)
                self.table.setUpdatesEnabled(False)
                self.table.clearContents()
                self.table.setRowCount(1)
                self.table.setColumnCount(1)
                self.sb_rows.setValue(1)
                self.sb_cols.setValue(1)
                it = QTableWidgetItem("")
                it.setTextAlignment(Qt.AlignCenter)
                it.setBackground(WHITE)
                self.table.setItem(0, 0, it)
            finally:
                self.table.setUpdatesEnabled(True)
                self.table.blockSignals(False)
            return

        est_cells = content_rows * content_cols
        truncated = est_cells > MAX_CELLS
        use_cols = content_cols
        use_rows = content_rows if not truncated else max(1, MAX_CELLS // max(1, use_cols))

        try:
            self.table.blockSignals(True)
            self.table.setUpdatesEnabled(False)

            self.table.clearContents()
            final_rows = max(use_rows, FIRST_DATA_ROW + 1)
            self.table.setRowCount(final_rows)
            self.table.setColumnCount(use_cols)
            self.sb_rows.setValue(final_rows)
            self.sb_cols.setValue(use_cols)

            row_idx = 0
            for row in sheet.getElementsByType(TableRow):
                if row_idx >= use_rows:
                    break
                rrep = int(row.getAttribute('numberrowsrepeated') or 1)

                # Build a visible template up to use_cols
                template = []
                col_idx = 0
                for cell in row.getElementsByType(TableCell):
                    crep = int(cell.getAttribute('numbercolumnsrepeated') or 1)
                    vis = min(crep, max(0, use_cols - col_idx))
                    if vis > 0:
                        text = _extract_text_from_cell(cell)
                        template.append((text, vis))
                    col_idx += crep
                    if col_idx >= use_cols:
                        break

                for _ in range(rrep):
                    if row_idx >= use_rows:
                        break
                    c = 0
                    for text, vis in template:
                        for _k in range(vis):
                            it = self.table.item(row_idx, c)
                            if it is None:
                                it = QTableWidgetItem("")
                                self.table.setItem(row_idx, c, it)
                            it.setTextAlignment(Qt.AlignCenter)
                            it.setText(text)
                            # раскрасим по новой логике
                            self.recolor_cell(it, row_idx, c)
                            c += 1
                            if c >= use_cols:
                                break
                        if c >= use_cols:
                            break
                    row_idx += 1

            # Fill remaining rows (если расширили вниз)
            for r in range(use_rows, final_rows):
                for c in range(use_cols):
                    it = self.table.item(r, c)
                    if it is None:
                        it = QTableWidgetItem("")
                        self.table.setItem(r, c, it)
                    it.setTextAlignment(Qt.AlignCenter)
                    it.setBackground(WHITE)

        finally:
            self.table.setUpdatesEnabled(True)
            self.table.blockSignals(False)

        # обновить панели из основной таблицы
        self._ensure_hdr_panel_cols()
        self._ensure_tol_panel_cols()
        self._sync_hdr_from_main()
        self._sync_tol_from_main()

        # прячем служебные строки
        if self.table.rowCount() > TOL_ROW:
            self.table.setRowHidden(TOL_ROW, True)
        for r in HEADER_ROWS:
            if r < self.table.rowCount():
                self.table.setRowHidden(r, True)

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
