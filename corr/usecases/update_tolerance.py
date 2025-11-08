from dataclasses import dataclass
from typing import Optional
from ..domain.types import TolSet
from ..domain.tolerance import parse_tol, format_with_opp, equal_tol
from ..usecases.recompute_metrics import recompute, Metrics


@dataclass
class TolUpdatePatch:
col: int
display_text: str
metrics: Metrics




def update_tolerance(grid, tolset: TolSet, col: int, new_text: str) -> TolUpdatePatch:
# вычисляем отображаемый текст с учётом ОПП
old_text = getattr(grid, 'orig_tol_texts', {}).get(col, "")
disp = format_with_opp(old_text, new_text)


# обновляем базу, если впервые
if old_text == "":
if not hasattr(grid, 'orig_tol_texts'):
grid.orig_tol_texts = {}
grid.orig_tol_texts[col] = new_text


tolset[col] = parse_tol(new_text)
metrics = recompute(grid, tolset)
return TolUpdatePatch(col=col, display_text=disp, metrics=metrics)