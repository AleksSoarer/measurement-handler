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
            

def get_cell_style(cell, ods_doc):
    """
    Возвращает dict с параметрами:
    {
        'bg': "#RRGGBB" или None,
        'text_color': "#RRGGBB" или None,
        'bold': True/False,
        'italic': True/False,
        'underline': True/False,
        'fontsize': число или None,
        ...
    }
    """
    # ---- 1. Фон (заливка) ----
    bg = get_bg_hex(cell, ods_doc)

    # ---- 2. Стиль текста ----
    sname = getattr(cell, "style_name", None)
    auto = getattr(ods_doc, "automatic_styles", None) or getattr(ods_doc, "automaticstyles", None)
    styles = getattr(ods_doc, "styles", None) or getattr(ods_doc, "styles_", None)
    st = None
    if sname:
        if auto and hasattr(auto, "get"):
            st = auto.get(sname)
        if not st and styles and hasattr(styles, "get"):
            st = styles.get(sname)

    # pyexcel-ezodf/odfpy: ищем текстовые свойства (font, underline, bold, color и т.д.)
    txt_props = None
    if st:
        for child in getattr(st, "childNodes", []):
            # Может быть <style:text-properties ...>
            qname = getattr(child, "qname", None)
            if qname and qname[1] == "text-properties":
                txt_props = child
                break

    result = {'bg': bg, 'text_color': None, 'bold': False, 'italic': False, 'underline': False, 'fontsize': None}

    if txt_props:
        color = txt_props.getAttribute("fo:color")
        if color:
            result['text_color'] = color
        if txt_props.getAttribute("fo:font-weight") == "bold":
            result['bold'] = True
        if txt_props.getAttribute("fo:font-style") == "italic":
            result['italic'] = True
        if txt_props.getAttribute("style:text-underline-style") not in (None, "none"):
            result['underline'] = True
        size = txt_props.getAttribute("fo:font-size")
        if size:
            try:
                result['fontsize'] = float(size.replace('pt',''))
            except Exception:
                result['fontsize'] = size

    return result

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
    Возвращает строку «#RRGGBB» или None, если у ячейки нет заливки,
    учитывая наследование стиля из строки, столбца и таблицы.
    """
    # 1. Стиль ячейки
    sname = getattr(cell, "style_name", None)
    hexclr = _get_style_bg(sname, ods_doc)
    if hexclr:
        return hexclr

    # 2. Стиль строки
    row = getattr(cell, "row", None)
    if row is not None:
        try:
            rowstyle = cell.table.rows[row].style_name
            hexclr = _get_style_bg(rowstyle, ods_doc)
            if hexclr:
                return hexclr
        except Exception:
            pass

    # 3. Стиль столбца
    col = getattr(cell, "column", None)
    if col is not None:
        try:
            colstyle = cell.table.columns[col].style_name
            hexclr = _get_style_bg(colstyle, ods_doc)
            if hexclr:
                return hexclr
        except Exception:
            pass

    # 4. Стиль по умолчанию для таблицы (default-cell-style)
    try:
        default_style = getattr(cell.table, "default_cell_style_name", None)
        hexclr = _get_style_bg(default_style, ods_doc)
        if hexclr:
            return hexclr
    except Exception:
        pass

    return None

def _get_style_bg(sname, ods_doc):
    if not sname:
        return None
    auto = getattr(ods_doc, "automatic_styles", None) or getattr(ods_doc, "automaticstyles", None)
    styles = getattr(ods_doc, "styles", None) or getattr(ods_doc, "styles_", None)
    st = None
    if auto and hasattr(auto, "get"):
        st = auto.get(sname)
    if not st and styles and hasattr(styles, "get"):
        st = styles.get(sname)
    if not st:
        return None
    # Старый ezodf/pyexcel-ezodf: dict-like .properties
    props = getattr(st, "properties", None)
    if isinstance(props, dict):
        return props.get("fo:background-color")
    # odfpy: getAttribute
    elif hasattr(st, "getAttribute"):
        return st.getAttribute("fo:background-color")
    return None



def clone_style_with_new_bg(ods_doc, orig_style_name, new_bg, cache):
    """
    Клонирует стиль orig_style_name, меняя только цвет фона.
    Сохраняет шрифты, границы и т.д.
    """
    key = (orig_style_name, new_bg)
    if key in cache:
        return cache[key]

    # Находим оригинальный стиль
    auto = getattr(ods_doc, "automatic_styles", None) or getattr(ods_doc, "automaticstyles", None)
    styles = getattr(ods_doc, "styles", None) or getattr(ods_doc, "styles_", None)

    orig_style = None
    if orig_style_name:
        if auto and hasattr(auto, "get"):
            orig_style = auto.get(str(orig_style_name))
        if not orig_style and styles and hasattr(styles, "get"):
            orig_style = styles.get(str(orig_style_name))
    # Если не нашли оригинал — создаём обычный новый
    if orig_style is None:
        style_name = ensure_style(ods_doc, new_bg, cache)
        cache[key] = style_name
        return style_name

    # Создаём новый стиль, копируя параметры
    from odf.style import Style, TableCellProperties
    new_name = f"{orig_style_name}_bg_{new_bg.lstrip('#')}"
    new_style = Style(name=new_name, family="table-cell")

    has_props = False
    for el in orig_style.childNodes:
        if getattr(el, "qname", (None, None))[1] == "table-cell-properties":
            has_props = True
            # Клонируем и меняем только backgroundcolor
            new_props = TableCellProperties()
            for attr in el.attributes.keys():
                if attr == "fo:background-color":
                    new_props.setAttribute(attr, new_bg)
                else:
                    new_props.setAttribute(attr, el.getAttribute(attr))
            new_style.addElement(new_props)
        else:
            # Копируем другие элементы (шрифт, границы и т.д.)
            new_style.addElement(el.cloneNode(True))
    # Если оригинал не содержит table-cell-properties, добавим свой
    if not has_props:
        new_style.addElement(TableCellProperties(backgroundcolor=new_bg))

    auto.addElement(new_style)
    cache[key] = new_name
    return new_name

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
                item = QTableWidgetItem(str(v) if v is not None else '')
                cell = self.sheet[data_start + r, fc + c]
                style_info = get_cell_style(cell, self.ods_doc)
                # Фон
                if style_info['bg']:
                    item.setBackground(QColor(style_info['bg']))
                # Цвет текста
                if style_info['text_color']:
                    item.setForeground(QColor(style_info['text_color']))
                # Жирный
                if style_info['bold']:
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                # Курсив
                if style_info['italic']:
                    font = item.font()
                    font.setItalic(True)
                    item.setFont(font)
                # Подчёркнутый
                if style_info['underline']:
                    font = item.font()
                    font.setUnderline(True)
                    item.setFont(font)
                # Размер
                if style_info['fontsize']:
                    font = item.font()
                    try:
                        font.setPointSize(int(float(style_info['fontsize'])))
                    except Exception:
                        pass
                    item.setFont(font)

                self.body_table.setItem(r, c, item)
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
                if gui_item is None:
                    continue
                hexclr = qcolor_hex(gui_item.background().color())
                ods_cell = self.sheet[ods_r, self.range_first_col + gui_c]
                prev_hex = get_bg_hex(ods_cell, self.ods_doc)
                
                # Только если у ячейки есть style_name и цвет реально менялся
                if ods_cell.style_name and prev_hex != hexclr:
                    style_name = clone_style_with_new_bg(self.ods_doc, ods_cell.style_name, hexclr, cache)
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