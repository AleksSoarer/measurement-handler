import re


_MARKERS = {"Y","N","Z","T","NM","Н","З","Т","НМ"}


BAD_TO_GOOD = (
("\u2212","-"), ("\u2013","-"), ("\u2014","-"),
("\u2012","-"), ("\u2010","-"),
("\u00A0",""), ("\u202F",""), ("\u2009",""), ("\u2007",""),
("\u2002",""), ("\u2003","")
)


_INT_RE = re.compile(r"^[0-9]+$")
_NUM_DOT_NUM_RE = re.compile(r"^[0-9]+\.[0-9]+$")




def normalize_num_str(s: str) -> str:
if s is None:
return ""
t = str(s).strip()
for bad, good in BAD_TO_GOOD:
t = t.replace(bad, good)
t = t.replace(" ", "").replace(",", ".")
return t




def try_parse_float(s: str):
if s is None:
return None
t = normalize_num_str(s)
if not t:
return None
u = t.upper()
if u in _MARKERS:
return None
try:
return float(t)
except ValueError:
return None




def fmt_serial(s: str) -> str:
f = try_parse_float(s)
if f is None:
return (s or "").strip()
i = int(round(f))
return str(i) if abs(f - i) < 1e-9 else str(f)