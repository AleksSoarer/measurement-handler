from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtGui import QPainter
from PyQt5.QtCore import QRect
from PyQt5.QtWidgets import QWidget, QTableWidget




def print_table_single_page(pdf_path: str, table: QTableWidget):
if table.rowCount() == 0 or table.columnCount() == 0:
raise RuntimeError("Таблица пуста")
vh = table.verticalHeader(); hh = table.horizontalHeader(); fw = table.frameWidth()*2
content_w = int(fw + vh.width() + sum(table.columnWidth(c) for c in range(table.columnCount())))
content_h = int(fw + hh.height() + sum(table.rowHeight(r) for r in range(table.rowCount())))


printer = QPrinter(QPrinter.HighResolution)
printer.setResolution(300)
printer.setOutputFormat(QPrinter.PdfFormat)
printer.setOutputFileName(pdf_path)
printer.setPaperSize(QPrinter.A4)
printer.setFullPage(True)
printer.setOrientation(QPrinter.Landscape if content_w >= content_h else QPrinter.Portrait)


painter = QPainter(printer)
target = printer.pageRect(QPrinter.DevicePixel)
sx = target.width() / float(content_w)
sy = target.height() / float(content_h)
scale = min(sx, sy)
view_w = max(1, int(content_w * scale))
view_h = max(1, int(content_h * scale))
offset_x = int((target.width() - view_w) / 2)
offset_y = int((target.height() - view_h) / 2)


painter.setViewport(offset_x, offset_y, view_w, view_h)
painter.setWindow(0, 0, content_w, content_h)
table.render(painter, flags=QWidget.DrawChildren)
painter.end()