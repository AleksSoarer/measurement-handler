from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.style import Style, TableCellProperties, TextProperties
from odf.table import Table, TableRow, TableCell
from odf.text import P


from ..shared.utils import try_parse_float, fmt_serial




def _extract_text(cell: TableCell) -> str:
parts = []
for p in cell.getElementsByType(P):
for node in getattr(p, 'childNodes', []):
data = getattr(node, 'data', None)
if data:
parts.append(str(data))
text = "".join(parts).strip()
if not text:
v = cell.getAttribute('value')
if v is not None:
text = str(v)
return text




def load_ods_matrix(path: str) -> list[list[str]]:
doc = load(path)
tables = doc.spreadsheet.getElementsByType(Table)
if not tables:
return [[]]
sheet = tables[0]
rows = []
max_cols = 0
for row in sheet.getElementsByType(TableRow):
rrep = int(row.getAttribute('numberrowsrepeated') or 1)
line = []
for cell in row.getElementsByType(TableCell):
crep = int(cell.getAttribute('numbercolumnsrepeated') or 1)
txt = _extract_text(cell)
line.extend([txt]*crep)
# подрезаем правые пустые
last = len(line)-1
while last >=0 and (str(line[last]).strip()==""):
last -= 1
line = line[:last+1]
for _ in range(rrep):
rows.append(list(line))
max_cols = max(max_cols, len(line))
out = [ (r + [""]*(max_cols - len(r))) for r in rows ]
return out