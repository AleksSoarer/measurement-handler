# -*- coding: utf-8 -*-
import sys
import re

# PyQt5
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QWidget,
    QVBoxLayout, QPushButton, QLineEdit, QTableWidget,
    QTableWidgetItem, QLabel, QSplitter
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

# Работа с ODS
import ezodf  # для opendoc/saveas
from odf.style import Style as EzStyle, TableCellProperties, TextProperties

# ────────────────────────────────────────────────────────────────────────
# Палитра для GUI
GOOD  = QColor("#C6EFCE")   # мягко-зелёный
BAD   = QColor("#FFC7CE")   # мягко-красный

# ────────────────────────────────────────────────────────────────────────
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ

def qcolor_hex(qc: QColor) -> str:
    return "#{:02x}{:02x}{:02x}".format(qc.red(), qc.green(), qc.blue())

def col2num(col: str) -> int:
    """ 'A' -> 0, 'AA' -> 26 """
    n = 0
    for c in col.upper():
        n = n * 26 + (ord(c) - 64)
    return n - 1

def parse_range(rng: str):
    """
    'A1:D10' -> (fc, fr, tc, tr) в 0-базисе.
    Пробелы/нижний регистр допускаются.
    """
    rng = rng.strip()
    m = re.fullmatch(r'\s*([A-Za-z]+)\s*(\d+)\s*:\s*([A-Za-z]+)\s*(\d+)\s*', rng)
    if not m:
        raise ValueError(f'Неверный диапазон: {rng!r}. Ожидается формат A1:D10.')
    fc, fr, tc, tr = m.groups()
    return col2num(fc), int(fr) - 1, col2num(tc), int(tr) - 1

def _parse_float(x):
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    if isinstance(x, str):
        s = x.strip().replace('\xa0', ' ').replace(',', '.')
        try:
            return float(s)
        except Exception:
            return None
    return None

def ods_number(cell):
    """
    Возвращает float для ячейки, даже если там формула:
    1) cell.value (если число),
    2) office:value (кэш формулы),
    3) парсинг текстовых узлов.
    """
    v = _parse_float(getattr(cell, 'value', None))
    if v is not None:
        return v

    el = getattr(cell, 'xmlnode', None) or getattr(cell, '_element', None) or getattr(cell, 'element', None)
    if el is not None and hasattr(el, 'getAttribute'):
        raw = el.getAttribute('office:value')
        v = _parse_float(raw)
        if v is not None:
            return v

    
    txt = None
    try:
        if el is not None:
            buf = []
            for n in getattr(el, 'childNodes', []):
                data = getattr(n, 'data', None)
                if isinstance(data, str):
                    buf.append(data)
            if buf:
                txt = ''.join(buf)
    except Exception:
        pass
    return _parse_float(txt)

def set_cell_style_name(cell, style_name: str):
    """
    Надёжно проставляет стиль ячейке:
    1) через свойство .style_name
    2) принудительно table:style-name в XML
    """
    try:
        cell.style_name = style_name
    except Exception:
        pass
    el = getattr(cell, 'xmlnode', None) or getattr(cell, '_element', None) or getattr(cell, 'element', None)
    if el is not None and hasattr(el, 'setAttribute'):
        try:
            el.setAttribute('table:style-name', style_name)
        except Exception:
            pass

# ────────────────────────────────────────────────────────────────────────
# Доступ к стилям документа

def _find_style_in_container(container, name: str):
    """Ищем стиль по name в контейнере (поддержка и .get, и перебора)."""
    if container is None:
        return None
    # Попытка через .get (pyexcel-ezodf)
    if hasattr(container, "get"):
        try:
            st = container.get(str(name))
            if st:
                return st
        except Exception:
            pass
    # Перебор (odfpy)
    try:
        for st in container.getElementsByType(EzStyle):
            if st.getAttribute('name') == name:
                return st
    except Exception:
        pass
    return None

def _get_style_by_name(doc, sname):
    if not sname:
        return None
    auto = getattr(doc, "automatic_styles", None) or getattr(doc, "automaticstyles", None)
    styles = getattr(doc, "styles", None) or getattr(doc, "styles_", None)
    st = _find_style_in_container(auto, str(sname))
    if st is None:
        st = _find_style_in_container(styles, str(sname))
    return st

def _get_style_bg(sname, ods_doc):
    st = _get_style_by_name(ods_doc, sname)
    if not st:
        return None
    # старый pyexcel-ezodf: dict-like
    props = getattr(st, "properties", None)
    if isinstance(props, dict):
        return props.get("fo:background-color")
    # odfpy: смотреть table-cell-properties
    if hasattr(st, "getAttribute"):
        # напрямую
        val = st.getAttribute("fo:background-color")
        if val:
            return val
        # через дочерние узлы
        try:
            for child in getattr(st, "childNodes", []):
                qname = getattr(child, "qname", None)
                if qname and qname[1] == "table-cell-properties":
                    v = child.getAttribute("fo:background-color")
                    if v:
                        return v
        except Exception:
            pass
    return None

def get_bg_hex(cell, ods_doc):
    """
    Возвращает цвет заливки конкретной ячейки с учётом наследования:
    ячейка -> строка -> столбец -> default-cell-style таблицы.
    """
    # 1) стиль ячейки
    sname = getattr(cell, "style_name", None)
    hexclr = _get_style_bg(sname, ods_doc)
    if hexclr:
        return hexclr

    # 2) стиль строки
    try:
        row_idx = getattr(cell, "row", None)
        if row_idx is not None:
            rowstyle = cell.table.rows[row_idx].style_name
            hexclr = _get_style_bg(rowstyle, ods_doc)
            if hexclr:
                return hexclr
    except Exception:
        pass

    # 3) стиль столбца
    try:
        col_idx = getattr(cell, "column", None)
        if col_idx is not None:
            colstyle = cell.table.columns[col_idx].style_name
            hexclr = _get_style_bg(colstyle, ods_doc)
            if hexclr:
                return hexclr
    except Exception:
        pass

    # 4) стиль по умолчанию таблицы
    try:
        default_style = getattr(cell.table, "default_cell_style_name", None)
        hexclr = _get_style_bg(default_style, ods_doc)
        if hexclr:
            return hexclr
    except Exception:
        pass
    return None

def get_cell_style(cell, ods_doc):
    """
    Возвращает dict:
      {'bg', 'text_color', 'bold', 'italic', 'underline', 'fontsize'}
    """
    result = {
        'bg': get_bg_hex(cell, ods_doc),
        'text_color': None,
        'bold': False,
        'italic': False,
        'underline': False,
        'fontsize': None,
    }
    sname = getattr(cell, "style_name", None)
    st = _get_style_by_name(ods_doc, sname)
    if not st:
        return result

    # Ищем text-properties
    try:
        for child in getattr(st, "childNodes", []):
            qname = getattr(child, "qname", None)
            if not qname:
                continue
            tag = qname[1]
            if tag == "text-properties":
                color = child.getAttribute("fo:color")
                if color:
                    result['text_color'] = color
                if child.getAttribute("fo:font-weight") == "bold":
                    result['bold'] = True
                if child.getAttribute("fo:font-style") == "italic":
                    result['italic'] = True
                if child.getAttribute("style:text-underline-style") not in (None, "none"):
                    result['underline'] = True
                size = child.getAttribute("fo:font-size")
                if size:
                    try:
                        result['fontsize'] = float(size.replace('pt',''))
                    except Exception:
                        result['fontsize'] = size
    except Exception:
        pass
    return result

# ────────────────────────────────────────────────────────────────────────
# СОЗДАНИЕ/КЛОНИРОВАНИЕ СТИЛЕЙ

def _auto_container(doc):
    return getattr(doc, "automatic_styles", None) or getattr(doc, "automaticstyles", None)

def _add_style_to_doc(doc, style_obj):
    auto = _auto_container(doc)
    if auto is not None:
        if hasattr(auto, "addElement"):
            auto.addElement(style_obj)
        else:
            # старые обёртки — list-like
            auto.append(style_obj)

def ensure_style(doc, hexclr: str, cache: dict) -> str:
    """
    Возвращает имя стиля с заданной заливкой, создаёт при необходимости.
    """
    if hexclr in cache:
        return cache[hexclr]
    sname = f"bg_{hexclr.lstrip('#')}"
    st = EzStyle(name=sname, family='table-cell')
    st.addElement(TableCellProperties(backgroundcolor=hexclr))
    _add_style_to_doc(doc, st)
    cache[hexclr] = sname
    return sname

def clone_style_with_new_bg(doc, orig_style_name: str, new_bg: str, cache: dict) -> str:
    """
    Клонирует существующий стиль ячейки, меняя ТОЛЬКО фон (backgroundcolor),
    сохраняя остальные свойства (границы, шрифт и т.п.).
    """
    key = (orig_style_name, new_bg)
    if key in cache:
        return cache[key]

    orig_style = _get_style_by_name(doc, orig_style_name)
    if orig_style is None:
        # исходник не нашли — просто создаём новый стиль с нужным фоном
        name = ensure_style(doc, new_bg, cache)
        cache[key] = name
        return name

    new_name = f"{orig_style_name}_bg_{new_bg.lstrip('#')}"
    new_style = EzStyle(name=new_name, family="table-cell")

    has_cell_props = False
    for el in getattr(orig_style, "childNodes", []):
        qname = getattr(el, "qname", (None, None))
        tag = qname[1] if qname else None
        if tag == "table-cell-properties":
            has_cell_props = True
            cp = TableCellProperties()
            for attr in el.attributes.keys():
                val = el.getAttribute(attr)
                if attr == "fo:background-color":
                    cp.setAttribute(attr, new_bg)
                else:
                    cp.setAttribute(attr, val)
            new_style.addElement(cp)
        else:
            new_style.addElement(el.cloneNode(True))

    if not has_cell_props:
        new_style.addElement(TableCellProperties(backgroundcolor=new_bg))

    _add_style_to_doc(doc, new_style)
    cache[key] = new_name
    return new_name

def ensure_style_with_text(doc, bg_color: str, text_color: str, cache: dict) -> str:
    """Стиль ячейки с заданными фоном и цветом текста."""
    key = (bg_color, text_color)
    if key in cache:
        return cache[key]
    style_name = f"bg_{bg_color.lstrip('#')}_text_{text_color.lstrip('#')}"
    st = EzStyle(name=style_name, family="table-cell")
    st.addElement(TableCellProperties(backgroundcolor=bg_color))
    st.addElement(TextProperties(color=text_color))
    _add_style_to_doc(doc, st)
    cache[key] = style_name
    return style_name

def clone_style_with_new_bg_text(doc, orig_style_name: str, new_bg: str, new_text: str, cache: dict) -> str:
    """Клонирует стиль, меняя фон и цвет текста, сохраняя остальное."""
    key = (orig_style_name, new_bg, new_text)
    if key in cache:
        return cache[key]

    orig_style = _get_style_by_name(doc, orig_style_name)
    if orig_style is None:
        # исходный стиль не нашли — создаём новый с нужными цветами
        name = ensure_style_with_text(doc, new_bg, new_text, cache)
        cache[key] = name
        return name

    new_style_name = f"{orig_style_name}_bg_{new_bg.lstrip('#')}_text_{new_text.lstrip('#')}"
    new_style = EzStyle(name=new_style_name, family="table-cell")

    has_cell_props = False
    has_text_props = False
    for el in getattr(orig_style, "childNodes", []):
        qname = getattr(el, "qname", (None, None))
        tag = qname[1] if qname else None
        if tag == "table-cell-properties":
            has_cell_props = True
            cp = TableCellProperties()
            for attr in el.attributes.keys():
                val = el.getAttribute(attr)
                if attr == "fo:background-color":
                    cp.setAttribute(attr, new_bg)
                else:
                    cp.setAttribute(attr, val)
            new_style.addElement(cp)
        elif tag == "text-properties":
            has_text_props = True
            tp = TextProperties()
            for attr in el.attributes.keys():
                val = el.getAttribute(attr)
                if attr == "fo:color":
                    tp.setAttribute(attr, new_text)
                else:
                    tp.setAttribute(attr, val)
            new_style.addElement(tp)
        else:
            new_style.addElement(el.cloneNode(True))

    if not has_cell_props:
        new_style.addElement(TableCellProperties(backgroundcolor=new_bg))
    if not has_text_props:
        new_style.addElement(TextProperties(color=new_text))

    _add_style_to_doc(doc, new_style)
    cache[key] = new_style_name
    return new_style_name

def clone_style_with_new_bg_clear_text(doc, orig_style_name: str, new_bg: str, cache: dict) -> str:
    """
    Клонирует стиль, меняя ТОЛЬКО фон и сбрасывая цвет текста (fo:color).
    Все прочие свойства сохраняются.
    """
    key = (orig_style_name, new_bg, '__clear_text_color__')
    if key in cache:
        return cache[key]

    orig = _get_style_by_name(doc, orig_style_name)
    if orig is None:
        name = ensure_style(doc, new_bg, cache)
        cache[key] = name
        return name

    new_name = f"{orig_style_name}_bg_{new_bg.lstrip('#')}_clrtext"
    new_style = EzStyle(name=new_name, family="table-cell")

    has_cell_props = False
    for el in getattr(orig, "childNodes", []):
        qname = getattr(el, "qname", (None, None))
        tag = qname[1] if qname else None

        if tag == "table-cell-properties":
            has_cell_props = True
            cp = TableCellProperties()
            for attr in el.attributes.keys():
                val = el.getAttribute(attr)
                if attr == "fo:background-color":
                    cp.setAttribute(attr, new_bg)
                else:
                    cp.setAttribute(attr, val)
            new_style.addElement(cp)

        elif tag == "text-properties":
            # копируем всё, кроме fo:color (сбрасываем цвет текста)
            tp = TextProperties()
            for attr in el.attributes.keys():
                if attr == "fo:color":
                    continue
                tp.setAttribute(attr, el.getAttribute(attr))
            new_style.addElement(tp)

        else:
            new_style.addElement(el.cloneNode(True))

    if not has_cell_props:
        new_style.addElement(TableCellProperties(backgroundcolor=new_bg))

    _add_style_to_doc(doc, new_style)
    cache[key] = new_name
    return new_name

# ────────────────────────────────────────────────────────────────────────
# ТАБЛИЦА ДЛЯ GUI

class ToleranceAwareTable(QTableWidget):
    def recheck_column(self, col: int, tol: float):
        for r in range(self.rowCount()):
            item = self.item(r, col)
            if item is None:
                continue
            text = item.text().strip()
            up = text.upper()

            # NM → чёрный фон + белый текст
            if up == "NM":
                item.setBackground(QColor("#000000"))
                item.setForeground(QColor("#FFFFFF"))
                continue

            # Y → зелёный фон
            if up == "Y":
                item.setBackground(GOOD)
                continue

            # Число: |value| <= tol → зелёный, иначе красный
            try:
                v = float(text)
            except ValueError:
                # другое текстовое значение — не трогаем фон в UI
                continue

            item.setBackground(GOOD if abs(v) <= tol else BAD)

# ────────────────────────────────────────────────────────────────────────
# ОСНОВНОЕ ОКНО

class ODSViewer(QMainWindow):
    NOMINAL_ROW = 4          # в файле (0-based)
    TOL_ROW     = 5

    def __init__(self):
        super().__init__()
        self.setWindowTitle('ODS viewer w/ tolerances')
        self.resize(1000, 700)

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
        self.head_table = QTableWidget(2, 0)
        self.head_table.setVerticalHeaderLabels(['Номинал','Допуск'])
        self.split.addWidget(self.head_table)

        # «тело» измерений M×N
        self.body_table = ToleranceAwareTable(0, 0)
        self.split.addWidget(self.body_table)

        # ── служебное
        self.sheet = None
        self.ods_doc = None
        self.current_path = ""
        self.range_first_col = 0

        # связка: если меняется допуск -> перепроверяем столбец
        self.head_table.cellChanged.connect(self._on_head_changed)

        self.btn_save = QPushButton('Сохранить как…')
        vbox.addWidget(self.btn_save)
        self.btn_save.clicked.connect(self.save_as)
        self.btn_save.setEnabled(False)          # активируем после открытия файла

    # -------- file open --------------------------------------------------
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, 'ODS', '', 'ODS (*.ods)')
        if not path:
            return
        self.ods_doc = ezodf.opendoc(path)       # сохраняем документ
        self.sheet   = self.ods_doc.sheets[0]
        self.current_path = path
        self.btn_save.setEnabled(True)
        self.show_range()                        # сразу отрисовываем

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
            self.head_table.setItem(0, c, QTableWidgetItem(str(nom_val) if nom_val is not None else ''))
            # допуск
            tol_val = self.sheet[self.TOL_ROW, fc + c].value
            self.head_table.setItem(1, c, QTableWidgetItem(str(tol_val) if tol_val is not None else ''))

        self.head_table.blockSignals(False)

        # ------ заполняем body_table ------------------------------------
        data_start = self.TOL_ROW + 1
        n_rows = tr - data_start + 1
        if n_rows < 0:
            n_rows = 0

        self.body_table.setRowCount(n_rows)
        self.body_table.setColumnCount(n_cols)

        for r in range(n_rows):
            for c in range(n_cols):
                v = self.sheet[data_start + r, fc + c].value
                item = QTableWidgetItem(str(v) if v is not None else '')
                cell = self.sheet[data_start + r, fc + c]
                style_info = get_cell_style(cell, self.ods_doc)

                # Фон из файла (в GUI), перед перекраской правилами
                if style_info['bg']:
                    try:
                        item.setBackground(QColor(style_info['bg']))
                    except Exception:
                        pass
                # Цвет текста
                if style_info['text_color']:
                    try:
                        item.setForeground(QColor(style_info['text_color']))
                    except Exception:
                        pass
                # Жирный/курсив/подчёркнутый/размер
                if style_info['bold']:
                    f = item.font(); f.setBold(True); item.setFont(f)
                if style_info['italic']:
                    f = item.font(); f.setItalic(True); item.setFont(f)
                if style_info['underline']:
                    f = item.font(); f.setUnderline(True); item.setFont(f)
                if style_info['fontsize']:
                    f = item.font()
                    try:
                        f.setPointSize(int(float(style_info['fontsize'])))
                    except Exception:
                        pass
                    item.setFont(f)

                self.body_table.setItem(r, c, item)

        # ------ первая глобальная проверка по допускам ------------------
        for c in range(n_cols):
            try:
                tol = float(self.head_table.item(1, c).text())
            except (TypeError, ValueError):
                tol = None
            if tol is not None:
                self.body_table.recheck_column(c, tol)

        # ------ второй проход: NM и Y для видимого диапазона -----------
        for r in range(self.body_table.rowCount()):
            for c in range(self.body_table.columnCount()):
                it = self.body_table.item(r, c)
                if not it:
                    continue
                up = it.text().strip().upper()
                if up == "NM":
                    it.setBackground(QColor("#000000"))
                    it.setForeground(QColor("#FFFFFF"))
                elif up == "Y":
                    it.setBackground(GOOD)

        self.range_first_col = fc      # пригодится при обратной записи допусков

    def _on_head_changed(self, row, col):
        if row != 1:
            return
        try:
            tol = float(self.head_table.item(1, col).text())
        except (TypeError, ValueError):
            return

        # пишем в таблицу (только видимый диапазон)
        ods_col = self.range_first_col + col
        self.sheet[self.TOL_ROW, ods_col].set_value(tol)

        # перекрашиваем столбец в GUI
        self.body_table.recheck_column(col, tol)

    # ────────────────────────────────────────────────────────────────────
    # СБРОС + ПОКРАСКА ВСЕГО ЛИСТА И СОХРАНЕНИЕ СТИЛЕЙ В ДОКУМЕНТ

    def push_colors_back(self):
        """
        1) Сбрасываем ВСЕ цвета (фон и цвет текста) в зоне данных листа.
        2) Снова красим по правилам:
           - NM -> чёрный фон + белый текст
           - Y  -> зелёный фон
           - число: |value| <= tol ? зелёный : красный
        """
        sheet = self.sheet
        doc   = self.ods_doc
        if sheet is None:
            return

        ncols = sheet.ncols()
        nrows = sheet.nrows()
        data_start = self.TOL_ROW + 1

        reset_cache = {}
        paint_cache = {}

        # заранее читаем допуски по всем столбцам (с учётом формул)
        tol_per_col = [ods_number(sheet[self.TOL_ROW, c]) for c in range(ncols)]

        # 1) СБРОС ЦВЕТОВ: фон -> белый, цвет текста -> по умолчанию
        for c in range(ncols):
            for r in range(data_start, nrows):
                cell = sheet[r, c]
                if getattr(cell, "style_name", None):
                    new_name = clone_style_with_new_bg_clear_text(doc, cell.style_name, "#FFFFFF", reset_cache)
                else:
                    new_name = ensure_style(doc, "#FFFFFF", reset_cache)
                set_cell_style_name(cell, new_name)

        # 2) ПОКРАСКА ПО ПРАВИЛАМ
        for c in range(ncols):
            tol = tol_per_col[c]
            for r in range(data_start, nrows):
                cell = sheet[r, c]
                val  = cell.value
                text = (str(val).strip() if val is not None else "")
                up   = text.upper()

                # NM -> чёрный фон + белый текст
                if up == "NM":
                    bg_hex, text_hex = "#000000", "#FFFFFF"
                    if getattr(cell, "style_name", None):
                        name = clone_style_with_new_bg_text(doc, cell.style_name, bg_hex, text_hex, paint_cache)
                    else:
                        name = ensure_style_with_text(doc, bg_hex, text_hex, paint_cache)
                    set_cell_style_name(cell, name)
                    continue

                # Y -> зелёный фон
                if up == "Y":
                    bg_hex = "#C6EFCE"
                    if getattr(cell, "style_name", None):
                        name = clone_style_with_new_bg(doc, cell.style_name, bg_hex, paint_cache)
                    else:
                        name = ensure_style(doc, bg_hex, paint_cache)
                    set_cell_style_name(cell, name)
                    continue

                # Число: сравнение по модулю с допуском
                vnum = ods_number(cell)
                if vnum is None or tol is None:
                    continue

                bg_hex = "#C6EFCE" if abs(vnum) <= tol else "#FFC7CE"
                if getattr(cell, "style_name", None):
                    name = clone_style_with_new_bg(doc, cell.style_name, bg_hex, paint_cache)
                else:
                    name = ensure_style(doc, bg_hex, paint_cache)
                set_cell_style_name(cell, name)

    def save_as(self):
        if not self.sheet:
            return

        # 1) синхронизируем допуски из шапки (текущий видимый диапазон) в строку TOL_ROW
        for c in range(self.head_table.columnCount()):
            it = self.head_table.item(1, c)
            if it is None:
                continue
            txt = (it.text() or "").strip()
            try:
                tol = float(txt)
            except Exception:
                continue
            self.sheet[self.TOL_ROW, self.range_first_col + c].set_value(tol)

        # 2) жёсткий сброс + покраска по правилам ВСЕГО листа
        self.push_colors_back()

        # 3) сохранить
        suggested = re.sub(r'\.ods$', '', self.current_path, flags=re.I) + '_checked.ods'
        path, _ = QFileDialog.getSaveFileName(self, 'Сохранить как…', suggested, 'ODS (*.ods)')
        if not path:
            return
        if not path.lower().endswith('.ods'):
            path += '.ods'

        self.ods_doc.saveas(path)

# ────────────────────────────────────────────────────────────────────────
# Точка входа

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = ODSViewer()
    win.show()
    sys.exit(app.exec_())
