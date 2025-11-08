from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem
from PyQt5.QtCore import Qt
from ..domain.types import Grid, CellMark
from ..shared.constants import HEADER_ROWS, NOMINAL_ROW, TOL_ROW, FIRST_DATA_ROW
from ..domain.palette import CellState, decide_colors
from ..shared.utils import try_parse_float


class TableBinder:
def __init__(self, table: QTableWidget):
self.table = table


def bind_grid(self, grid: Grid):
t = self.table
t.blockSignals(True)
try:
t.clearContents()
t.setRowCount(len(grid.rows))
t.setColumnCount(grid.cols)
for r, row in enumerate(grid.rows):
for c in range(grid.cols):
txt = row.cells[c].text
it = t.item(r, c) or QTableWidgetItem("")
it.setTextAlignment(Qt.AlignCenter)
it.setText(txt)
is_service = r in HEADER_ROWS or r in (NOMINAL_ROW, TOL_ROW)
val = try_parse_float(txt)
mark = row.cells[c].mark
state = CellState(
text=txt,
mark=mark,
is_service=is_service,
is_col0=(c==0),
is_nm=(txt.strip().upper() in ("NM","НМ")),
is_num=(val is not None and r>=FIRST_DATA_ROW and c>0),
in_tol=True # окрасим позже патчем, здесь нейтрально
)
bg, fg = decide_colors(state)
it.setBackground(bg); it.setForeground(fg)
t.setItem(r, c, it)
# скрываем кол.0 в основной таблице
if grid.cols>0:
t.setColumnHidden(0, True)
finally:
t.blockSignals(False)


def apply_oos_coloring(self, grid: Grid, tolset):
# быстрый проход: только данные ниже FIRST_DATA_ROW
for r in range(FIRST_DATA_ROW, len(grid.rows)):
for c in range(1, grid.cols):
it = self.table.item(r, c)
txt = (it.text() or "").strip()
is_nm = txt.upper() in ("NM","НМ")
val = try_parse_float(txt)
in_tol = True
tol = tolset[c] if c < len(tolset) else None
if tol is not None and val is not None:
if hasattr(tol, 'value'):
in_tol = abs(val) <= tol.value
else:
in_tol = tol.lo <= val <= tol.hi
state = CellState(txt, CellMark.NONE, False, False, is_nm, val is not None, in_tol)
bg, fg = decide_colors(state)
it.setBackground(bg); it.setForeground(fg)