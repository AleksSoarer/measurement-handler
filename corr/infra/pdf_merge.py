from pypdf import PdfReader, PdfWriter


def merge_pdfs(out_path: str, *pdf_paths: str):
w = PdfWriter()
for p in pdf_paths:
if not p:
continue
r = PdfReader(p)
if getattr(r, 'is_encrypted', False):
try:
r.decrypt("")
except Exception:
continue
for page in r.pages:
w.add_page(page)
with open(out_path, 'wb') as f:
w.write(f)