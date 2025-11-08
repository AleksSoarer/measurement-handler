from .types import CellMark
from ..shared.constants import GREEN, RED, BLUE, WHITE, BLACK, TEXT


class CellState:
# простая контейнерная структура для решения окраски
def __init__(self, text: str, mark: CellMark, is_service: bool,
is_col0: bool, is_nm: bool, is_num: bool, in_tol: bool):
self.text=text; self.mark=mark; self.is_service=is_service
self.is_col0=is_col0; self.is_nm=is_nm; self.is_num=is_num; self.in_tol=in_tol




def decide_colors(state: CellState):
if state.is_service:
return WHITE, TEXT
if state.is_col0:
return WHITE, TEXT
if state.is_nm:
return BLACK, WHITE
if state.mark in (CellMark.N, CellMark.Z, CellMark.T):
return RED, TEXT
if state.mark == CellMark.Y:
return GREEN, TEXT
if state.is_num:
return (BLUE if state.in_tol else RED), TEXT
if any(ch.isalpha() for ch in state.text or ""):
return RED, TEXT
if any(ch.isdigit() for ch in state.text or ""):
return GREEN, TEXT
return WHITE, TEXT