from dataclasses import dataclass
from typing import List
from ..domain.types import Grid, TolSet
from ..domain.rules import count_oos, count_totals


@dataclass
class Metrics:
oos_per_col: List[int]
total: int
good: int
bad: int




def recompute(grid: Grid, tolset: TolSet) -> Metrics:
oos = count_oos(grid, tolset)
total, good, bad = count_totals(grid, tolset)
return Metrics(oos, total, good, bad)