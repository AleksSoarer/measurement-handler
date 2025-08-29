import sys
import re
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QSpinBox, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog,
    QMessageBox
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
BLUE  = QColor("#9DC3E6")   # soft blue
WHITE = QColor("#FFFFFF")

# Global numeric rule
DELTA = 0.5  # threshold for "number > DELTA -> BLUE"

# Header rows (0-based indices)
NOMINAL_ROW = 4        # 5th row shows "Nominal size"
TOL_ROW     = 5        # 6th row shows "Tolerance"
FIRST_DATA_ROW = 6     # data start row

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
        self.setWindowTitle("Мини-редактор ODS (цвета сохраняются, допуски)")
        self.resize(1000, 700)

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

        # Table
        self.table = QTableWidget(0, 0, self)
        self.table.cellChanged.connect(self.on_cell_changed)
        root.addWidget(self.table)

        # Initial table
        self.build_table()

    # ---- Coloring rules ----
    def _get_nom_tol(self, col):
        """Return (nominal, tol) for given column or (None, None). Skip col=0."""
        if col == 0:
            return None, None
        nom_it = self.table.item(NOMINAL_ROW, col)
        tol_it = self.table.item(TOL_ROW, col)
        nom = try_parse_float(nom_it.text()) if nom_it else None
        tol = try_parse_float(tol_it.text()) if tol_it else None
        return nom, tol

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

        # Priority 0: hard markers
        if up == "NM":
            it.setBackground(RED)
            return
        if up == "Y":
            it.setBackground(GREEN)
            return

        f = try_parse_float(text)

        # Priority 1: tolerance logic for data rows and columns > 0
        if (row >= FIRST_DATA_ROW) and (col > 0):
            nom, tol = self._get_nom_tol(col)
            if (nom is not None) and (tol is not None) and (f is not None):
                if abs(f - nom) <= tol:
                    it.setBackground(BLUE)   # in tolerance band
                else:
                    it.setBackground(RED)    # out of tolerance
                return

        # Priority 2: legacy DELTA rule
        if f is not None and f > DELTA:
            it.setBackground(BLUE)
            return

        # Priority 3: generic
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
        """Repaint entire column after nominal/tolerance change."""
        if col <= 0:
            return
        nom, tol = self._get_nom_tol(col)
        if nom is None or tol is None:
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

    # ---- UI callbacks ----
    def build_table(self):
        cols = max(self.sb_cols.value(), 1)
        # ensure we have at least header rows
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
        finally:
            self.table.blockSignals(False)

    def on_cell_changed(self, row, col):
        # If nominal/tolerance edited, repaint entire column
        if row in (NOMINAL_ROW, TOL_ROW):
            self.recheck_column(col)
            return
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

        # 1) get content bounds instead of expanding repeated empties
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

        # 2) cap by MAX_CELLS
        est_cells = content_rows * content_cols
        truncated = est_cells > MAX_CELLS
        use_cols = content_cols
        use_rows = content_rows if not truncated else max(1, MAX_CELLS // max(1, use_cols))

        try:
            self.table.blockSignals(True)
            self.table.setUpdatesEnabled(False)

            self.table.clearContents()
            # Also ensure we still have header rows visible even if file is tiny
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
                            # we ignore external styles and recolor by our rules
                            self.recolor_cell(it, row_idx, c)
                            c += 1
                            if c >= use_cols:
                                break
                        if c >= use_cols:
                            break
                    row_idx += 1

            # Fill remaining rows (if we extended to ensure header presence)
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
