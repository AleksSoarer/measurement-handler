import sys, ezodf, re
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QWidget,
    QVBoxLayout, QPushButton, QLineEdit, QTableWidget,
    QTableWidgetItem, QLabel, QSplitter
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

GOOD  = QColor("#C6EFCE")   # мягко-зелёный
BAD   = QColor("#FFC7CE")   # мягко-красный

def recheck_column(self, col, nominal, tol):
    lo, hi = nominal - tol, nominal + tol
    for r in range(self.rowCount()):
        it = self.item(r, col)
        if it is None:
            continue
        try:
            v = float(it.text())
        except ValueError:
            # не число - вернём «родной» фон
            it.setBackground(it.data(Qt.UserRole) or Qt.white)
            continue

        if lo <= v <= hi:
            it.setBackground(GOOD)
        else:
            it.setBackground(BAD)

def qcolor_hex(qc):
    return "#{:02x}{:02x}{:02x}".format(qc.red(), qc.green(), qc.blue())

def ensure_style(doc, hexclr, cache):
    """
    Возвращает имя стиля с заданной заливкой, добавляя его в
    automatic_styles, если такого ещё нет.
    """
    if hexclr not in cache:
        sname = f"bg_{hexclr.lstrip('#')}"
        st = ezodf.Style(name=sname, family="table-cell")
        st.set_property("fo:background-color", hexclr)       # :contentReference[oaicite:1]{index=1}
        doc.automatic_styles.append(st)
        cache[hexclr] = sname
    return cache[hexclr]

def push_colors_back(self):
    cache = {}
    body_start = self.TOL_ROW + 1
    for gui_row in range(self.body_table.rowCount()):
        ods_row = body_start + gui_row
        for col in range(self.body_table.columnCount()):
            gui_it = self.body_table.item(gui_row, col)
            hexclr  = qcolor_hex(gui_it.background().color())
            ods_cell = self.sheet[ods_row, self.range_first_col + col]
            style_name = ensure_style(self.ods_doc, hexclr, cache)
            ods_cell.style_name = style_name
            



# ---------- вспомогалка -------------------------------------------------
def col2num(col: str) -> int:                 # 'A' -> 0, 'AA' -> 26
    n = 0
    for c in col:
        n = n * 26 + (ord(c) - 64)
    return n - 1

def parse_range(rng: str):
    m = re.fullmatch(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', rng)
    if not m:
        raise ValueError(f'Неверный диапазон: {rng}')
    fc, fr, tc, tr = m.groups()
    return col2num(fc), int(fr) - 1, col2num(tc), int(tr) - 1

def get_bg_hex(cell, ods_doc):
    """
    Возвращает строку «#RRGGBB» или None, если у ячейки нет заливки.
    """
    sname = cell.style_name
    if not sname:
        return None

    # automatic_styles в приоритете, иначе ищем в styles
    st = ods_doc.automatic_styles.get(sname) or ods_doc.styles.get(sname)
    if not st:
        return None
    return st.properties.get('fo:background-color')          # спецификация ODF 1.1 :contentReference[oaicite:0]{index=0}


# ---------- таблица, умеющая подсвечивать -------------------------------
class ToleranceAwareTable(QTableWidget):
    def recheck_column(self, col: int, nominal: float, tol: float):
        low, high = nominal - tol, nominal + tol
        for r in range(self.rowCount()):
            item = self.item(r, col)
            if item is None:
                continue
            try:
                v = float(item.text())
            except ValueError:                # текст – пропускаем
                item.setBackground(Qt.white)
                continue
            item.setBackground(Qt.green if low <= v <= high else Qt.red)

# ---------- основное окно -----------------------------------------------
class ODSViewer(QMainWindow):
    NOMINAL_ROW = 4          # в файле (0-based)
    TOL_ROW     = 5

    def __init__(self):
        super().__init__()
        self.setWindowTitle('ODS viewer w/ tolerances')
        self.resize(900, 700)

        # ── UI -----------------------------------------------------------
        main = QWidget(self); self.setCentralWidget(main)
        vbox = QVBoxLayout(main)

        self.btn_open  = QPushButton('Открыть ODS'); vbox.addWidget(self.btn_open)
        self.btn_open.clicked.connect(self.open_file)

        self.range_in  = QLineEdit('A1:D10')
        vbox.addWidget(QLabel('Диапазон (опц.):')); vbox.addWidget(self.range_in)

        self.btn_show  = QPushButton('Показать диапазон'); vbox.addWidget(self.btn_show)
        self.btn_show.clicked.connect(self.show_range)

        self.split     = QSplitter(Qt.Vertical); vbox.addWidget(self.split, 1)

        # «шапка» 2×N
        self.head_table = QTableWidget(2, 0); self.head_table.setVerticalHeaderLabels(['Номинал','Допуск'])
        self.split.addWidget(self.head_table)

        # «тело» измерений M×N
        self.body_table = ToleranceAwareTable(0, 0); self.split.addWidget(self.body_table)

        # ── служебное
        self.sheet = None

        # связка: если меняется допуск -> перепроверяем столбец
        self.head_table.cellChanged.connect(self._on_head_changed)

    # -------- file open --------------------------------------------------
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, 'ODS', '', 'ODS (*.ods)')
        if not path: return
        self.sheet = ezodf.opendoc(path).sheets[0]
        self.show_range()                     # сразу отрисовываем

    # -------- головная проверка изменения допусков ----------------------
    def _on_head_changed(self, row, col):
        if row != 1:                 # 0 – номинал, 1 – допуск
            return
        try:
            nominal = float(self.head_table.item(0, col).text())
            tol     = float(self.head_table.item(1, col).text())
        except (TypeError, ValueError):
            return
        self.body_table.recheck_column(col, nominal, tol)

    # -------- отображение диапазона или всей таблицы --------------------
    def show_range(self):
        if not self.sheet:
            return
        # ------ определяем границы --------------------------------------
        rng = self.range_in.text().strip()
        if rng:
            fc, fr, tc, tr = parse_range(rng)
        else:
            fc, fr, tc, tr = 0, 0, self.sheet.ncols()-1, self.sheet.nrows()-1

        # ------ кол-во столбцов -----------------------------------------
        n_cols = tc - fc + 1
        # ------ заполняем head_table ------------------------------------
        self.head_table.setColumnCount(n_cols)

        for c in range(n_cols):
            # номинал
            nom_val = self.sheet[self.NOMINAL_ROW, fc + c].value
            self.head_table.setItem(0, c, QTableWidgetItem(str(nom_val) if nom_val else ''))
            # допуск
            tol_val = self.sheet[self.TOL_ROW, fc + c].value
            self.head_table.setItem(1, c, QTableWidgetItem(str(tol_val) if tol_val else ''))

        # ------ заполняем body_table ------------------------------------
        data_start = self.TOL_ROW + 1
        n_rows = tr - data_start + 1
        self.body_table.setRowCount(n_rows); self.body_table.setColumnCount(n_cols)

        for r in range(n_rows):
            for c in range(n_cols):
                v = self.sheet[data_start + r, fc + c].value
                self.body_table.setItem(r, c, QTableWidgetItem(str(v) if v is not None else ''))

        # ------ первая глобальная проверка ------------------------------
        for c in range(n_cols):
            try:
                nominal = float(self.head_table.item(0, c).text())
                tol     = float(self.head_table.item(1, c).text())
            except (TypeError, ValueError):
                continue
            self.body_table.recheck_column(c, nominal, tol)

# ---------- main ---------------------------------------------------------
if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = ODSViewer(); win.show()
    sys.exit(app.exec_())