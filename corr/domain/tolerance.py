import re
from .types import TolKind, ScalarTol, RangeTol, Tol
from typing import Optional, Tuple
from ..shared.utils import normalize_num_str


_NUM_RE = r"[-−]?\d+(?:[.,]\d+)?"
_NUM_ONLY_RE = re.compile(r"^\d+(?:[.,]\d+)?$")
_NUM_SLASH_RE = re.compile(rf"^\s*{_NUM_RE}\s*[\\/]\s*{_NUM_RE}\s*$")


_OPP_DECOR_RE = re.compile(
r"^\s*([0-9]+(?:[.,][0-9]+)?)\s*\(\s*ОПП\s*([0-9]+(?:[.,][0-9]+)?)\s*\)\s*$",
re.IGNORECASE,
)
_OPP_SLASH_DECOR_RE = re.compile(
r"^\s*((?:" + _NUM_RE + r")\s*[\\/]\s*(?:" + _NUM_RE + r"))\s*\(\s*ОПП\s*((?:" + _NUM_RE + r")\s*[\\/]\s*(?:" + _NUM_RE + r"))\s*\)\s*$",
re.IGNORECASE,
)




def _to_float(s: str) -> float:
return float(normalize_num_str(s))




def parse_slash_pair(s: str) -> Tuple[float, float]:
s = normalize_num_str(s)
m = re.fullmatch(rf"({_NUM_RE})[\\/]({_NUM_RE})", s)
if not m:
raise ValueError("bad slash tol")
a,b = _to_float(m.group(1)), _to_float(m.group(2))
return (a,b) if a <= b else (b,a)




def detect_kind(raw: str) -> Tuple[TolKind, str]:
s = (raw or "").strip()
if not s:
return (TolKind.SYMBOLIC, "")
if _OPP_SLASH_DECOR_RE.fullmatch(s):
# возвратим new часть как текущее значение
new = _OPP_SLASH_DECOR_RE.fullmatch(s).group(2)
return (TolKind.RANGE, new)
if _NUM_SLASH_RE.fullmatch(s):
return (TolKind.RANGE, s)
if _OPP_DECOR_RE.fullmatch(s):
new = _OPP_DECOR_RE.fullmatch(s).group(2)
return (TolKind.SCALAR, new)
if _NUM_ONLY_RE.fullmatch(s):
return (TolKind.SCALAR, s)
return (TolKind.SYMBOLIC, s)




def parse_tol(raw: str) -> Optional[Tol]:
kind, val = detect_kind(raw)
if kind == TolKind.SCALAR:
return ScalarTol(_to_float(val))
if kind == TolKind.RANGE:
lo, hi = parse_slash_pair(val)
return RangeTol(lo, hi)
return None # символика/пусто → не считаем




def equal_tol(a: Optional[Tol], b: Optional[Tol]) -> bool:
if a is None or b is None:
return a is None and b is None
if type(a) is not type(b):
return False
if isinstance(a, ScalarTol):
return abs(a.value - b.value) < 1e-12
return abs(a.lo - b.lo) < 1e-12 and abs(a.hi - b.hi) < 1e-12




def check_delta(delta: float, tol: Tol) -> bool:
if isinstance(tol, ScalarTol):
return abs(delta) <= tol.value
return tol.lo <= delta <= tol.hi




def format_with_opp(old_raw: str, new_raw: str) -> str:
old = parse_tol(old_raw)
new = parse_tol(new_raw)
if equal_tol(old, new):
return old_raw or new_raw
return f"{old_raw} (ОПП {new_raw})" if old_raw else new_raw