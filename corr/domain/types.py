from dataclasses import dataclass
from enum import Enum, auto
from typing import Optional, List, Tuple, Union


class TolKind(Enum):
SCALAR = auto()
RANGE = auto() # a/b
SYMBOLIC = auto()


@dataclass
class ScalarTol:
value: float


@dataclass
class RangeTol:
lo: float
hi: float


Tol = Union[ScalarTol, RangeTol]


class CellMark(Enum):
NONE=auto(); Y=auto(); N=auto(); Z=auto(); T=auto(); NM=auto()


@dataclass
class Cell:
text: str = ""
value: Optional[float] = None
mark: CellMark = CellMark.NONE


@dataclass
class Row:
cells: List[Cell]
is_service: bool = False


@dataclass
class Grid:
rows: List[Row]
cols: int


# Tolerances на колонку: индекс -> Tol или None
TolSet = List[Optional[Tol]]