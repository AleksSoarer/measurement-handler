from dataclasses import dataclass, field
from typing import List, Optional, Dict, Tuple

@dataclass
class Grid:
    # прямоугольная матрица строк как базовый универсальный формат
    rows: int
    cols: int
    data: List[List[str]]  # data[r][c] -> текст

@dataclass
class ToleranceSet:
    # допуска по колонкам, None если нечисловой/неанализируемый
    scalar: Dict[int, Optional[float]] = field(default_factory=dict)           # c -> tol (abs)
    slash: Dict[int, Optional[Tuple[float, float]]] = field(default_factory=dict)  # c -> (lo, hi)
    nonnumeric_cols: set = field(default_factory=set)  # колонки с символикой

@dataclass
class Metrics:
    total_parts: int = 0
    good_parts: int = 0
    bad_per_col: Dict[int, int] = field(default_factory=dict)  # c -> count