from dataclasses import dataclass
from typing import Optional
from ..domain.types import Grid, Row, Cell, TolSet
from ..shared.constants import HEADER_ROWS, TOL_ROW, FIRST_DATA_ROW
from ..domain.tolerance import parse_tol
from ..usecases.recompute_metrics import recompute, Metrics


@dataclass
class LoadResult:
grid: Grid
tolset: TolSet
metrics: Metrics




def _build_grid_from_matrix(matrix: list[list[str]]) -> Grid:
rows = []
for r, line in enumerate(matrix):
cells = [Cell(text=str(v or "")) for v in line]
is_service = r in HEADER_ROWS or r in (TOL_ROW,)
rows.append(Row(cells=cells, is_service=is_service))
cols = len(matrix[0]) if matrix else 0
return Grid(rows=rows, cols=cols)




def load_from_matrix(matrix: list[list[str]]) -> LoadResult:
grid = _build_grid_from_matrix(matrix)
tolset: TolSet = [None] * grid.cols
if grid.cols:
for c in range(grid.cols):
if c == 0:
continue
txt = grid.rows[TOL_ROW].cells[c].text if TOL_ROW < len(grid.rows) else ""
tolset[c] = parse_tol(txt)
metrics = recompute(grid, tolset)
return LoadResult(grid, tolset, metrics)