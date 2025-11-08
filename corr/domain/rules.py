from typing import List, Optional
from .types import Grid, TolSet, RangeTol, ScalarTol, CellMark
from ..shared.utils import try_parse_float
from ..shared.constants import FIRST_DATA_ROW


BAD_MARKS = {CellMark.N, CellMark.Z, CellMark.T}




def is_row_defective(grid: Grid, r: int, tolset: TolSet) -> bool:
# нет серийника → строка вне учёта и не брак
sn = (grid.rows[r].cells[0].text or "").strip()
if not sn:
return False
cols = grid.cols
if cols <= 1:
return True


has_value = False
for c in range(1, cols):
cell = grid.rows[r].cells[c]
txt = (cell.text or "").strip()
if txt:
has_value = True
# марки
if cell.mark in BAD_MARKS:
return True
f = try_parse_float(txt)
if f is None:
continue
tol = tolset[c] if c < len(tolset) else None
if tol is None:
continue
ok = (abs(f) <= tol.value) if isinstance(tol, ScalarTol) else (tol.lo <= f <= tol.hi)
if not ok:
return True
return not has_value




def count_oos(grid: Grid, tolset: TolSet) -> List[int]:
cols = grid.cols
rows = len(grid.rows)
out = [0]*cols
for c in range(cols):
if c == 0:
continue
tol = tolset[c] if c < len(tolset) else None
cnt = 0
for r in range(FIRST_DATA_ROW, rows):
sn = (grid.rows[r].cells[0].text or "").strip()
if not sn:
continue
cell = grid.rows[r].cells[c]
txt = (cell.text or "").strip()
if not txt:
continue
if cell.mark in BAD_MARKS:
cnt += 1; continue
f = try_parse_float(txt)
if f is None:
continue
if tol is None: # символика — пропускаем
continue
ok = (abs(f) <= tol.value) if isinstance(tol, ScalarTol) else (tol.lo <= f <= tol.hi)
if not ok:
cnt += 1
out[c] = cnt
return out




def count_totals(grid: Grid, tolset: TolSet):
rows = len(grid.rows)
total = good = 0
for r in range(FIRST_DATA_ROW, rows):
sn = (grid.rows[r].cells[0].text or "").strip()
if not sn:
continue
total += 1
if not is_row_defective(grid, r, tolset):
good += 1
bad = total - good
return total, good, bad