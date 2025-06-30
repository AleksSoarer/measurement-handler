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


def qcolor_hex(qc):
    return "#{:02x}{:02x}{:02x}".format(qc.red(), qc.green(), qc.blue())

from typing import Callable

try:                                     # 1. самый старый ezodf (<=0.3.0)
    from ezodf import Style as EzStyle   # type: ignore
    def _mk_cell_props(color: str, name: str):
        st = EzStyle(name=name, family='table-cell')
        st.set_property('fo:background-color', color)
        return st

except ImportError:
    try:                                 # 2. «средний» ezodf (есть ezodf.style)
        from ezodf.style import Style as EzStyle  # type: ignore
        def _mk_cell_props(color: str, name: str):
            st = EzStyle(name=name, family='table-cell')
            st.set_property('fo:background-color', color)
            return st

    except ImportError:                  # 3. современный путь — odfpy
        from odf.style import Style as EzStyle, TableCellProperties
        def _mk_cell_props(color: str, name: str):
            st = EzStyle(name=name, family='table-cell')
            st.addElement(TableCellProperties(backgroundcolor=color))
            return st

def ensure_style(doc, hexclr: str, cache: dict[str, str]) -> str:
    """
    Гарантированно возвращает имя стиля с нужной заливкой,
    создавая его при необходимости. Совместимо со старыми ezodf
    (append) и pyexcel-ezodf / odfpy (addElement).
    """
    if hexclr in cache:
        return cache[hexclr]

    sname = f"bg_{hexclr.lstrip('#')}"
    st = _mk_cell_props(hexclr, sname)

    # ── выбираем коллекцию автоматических стилей ─────────────────────
    auto = getattr(doc, "automatic_styles", None) or getattr(doc, "automaticstyles", None)
    if auto is None:
        # крайне редкий случай – у документа нет ни одного блока стилей
        return sname   # просто пропускаем, без выброса исключения

    # ── добавляем стиль подходящим методом ───────────────────────────
    if hasattr(auto, "addElement"):        # odfpy / pyexcel-ezodf
        auto.addElement(st)
    else:                                  # старый ezodf: обычный list-like
        auto.append(st)
    # ─────────────────────────────────────────────────────────────────

    cache[hexclr] = sname
    return sname
            



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
    def recheck_column(self, col: int, tol: float):
        for r in range(self.rowCount()):
            item = self.item(r, col)
            if item is None:
                continue
            try:
                v = float(item.text())
            except ValueError:              # в ячейке не число
                item.setBackground(Qt.white)
                continue

            # |Δ| <= tol  → зелёный   |Δ| > tol → красный
            item.setBackground(GOOD if abs(v) <= tol else BAD)

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
        self.range_first_col = 0

        # связка: если меняется допуск -> перепроверяем столбец
        self.head_table.cellChanged.connect(self._on_head_changed)

        self.btn_save = QPushButton('Сохранить как…')
        vbox.addWidget(self.btn_save)
        self.btn_save.clicked.connect(self.save_as)
        self.btn_save.setEnabled(False)          # активируем только после открытия файла

    # -------- file open --------------------------------------------------
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, 'ODS', '', 'ODS (*.ods)')
        if not path:
            return
        self.ods_doc = ezodf.opendoc(path)       # <── сохраняем документ
        self.sheet   = self.ods_doc.sheets[0]
        self.current_path = path
        self.btn_save.setEnabled(True)
        self.show_range()                   # сразу отрисовываем



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

        self.range_first_col = fc

        

        # на время массового заполнения «шапки» отключаем сигнал
        self.head_table.blockSignals(True)

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
        
        self.head_table.blockSignals(False) 

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
                tol = float(self.head_table.item(1, c).text())
            except (TypeError, ValueError):
                continue
            self.body_table.recheck_column(c, tol)

        self.range_first_col = fc      # пригодится при обратной записи

    def _on_head_changed(self, row, col):
        if row != 1:
            return
        try:
            tol = float(self.head_table.item(1, col).text())
        except (TypeError, ValueError):
            return

        # пишем в таблицу
        ods_col = self.range_first_col + col
        self.sheet[self.TOL_ROW, ods_col].set_value(tol)

        # перекрашиваем столбец
        self.body_table.recheck_column(col, tol)

    def push_colors_back(self):
        cache = {}
        body_start = self.TOL_ROW + 1
        for gui_r in range(self.body_table.rowCount()):
            ods_r = body_start + gui_r
            for gui_c in range(self.body_table.columnCount()):
                gui_item = self.body_table.item(gui_r, gui_c)
                hexclr   = qcolor_hex(gui_item.background().color())
                ods_cell = self.sheet[ods_r, self.range_first_col + gui_c]
                style_name = ensure_style(self.ods_doc, hexclr, cache)
                ods_cell.style_name = style_name

    def save_as(self):
        if not self.sheet:
            return

        # 1) гарантируем, что все допуски из шапки попали в лист
        for c in range(self.head_table.columnCount()):
            try:
                tol = float(self.head_table.item(1, c).text())
                self.sheet[self.TOL_ROW, self.range_first_col + c].set_value(tol)
            except (TypeError, ValueError):
                pass

        # 2) переносим цвета
        self.push_colors_back()

        # 3) спрашиваем имя и сохраняем
        suggested = self.current_path.replace('.ods', '_checked.ods')
        path, _ = QFileDialog.getSaveFileName(self, 'Сохранить как…', suggested, 'ODS (*.ods)')
        if path:
            self.ods_doc.saveas(path)

# ---------- main ---------------------------------------------------------
if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = ODSViewer(); win.show()
    sys.exit(app.exec_())