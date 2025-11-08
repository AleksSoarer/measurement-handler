from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTableWidget
from PyQt5.QtCore import Qt
from ..shared.constants import UI_FONT_PT
from .table_binder import TableBinder


class MainWindow(QWidget):
def __init__(self, presenter_cls):
super().__init__()
self.setWindowTitle("Контроль допусков — новая архитектура")
root = QVBoxLayout(self)
bar = QHBoxLayout()
btn_open_xlsx = QPushButton("Открыть .xlsx"); btn_open_xlsx.clicked.connect(self._open_xlsx)
btn_open_ods = QPushButton("Открыть .ods"); btn_open_ods.clicked.connect(self._open_ods)
btn_save_xlsx = QPushButton("Сохранить .xlsx"); btn_save_xlsx.clicked.connect(self._save_xlsx)
btn_pdf = QPushButton("PDF: таблица"); btn_pdf.clicked.connect(self._save_pdf)
for b in (btn_open_xlsx, btn_open_ods, btn_save_xlsx, btn_pdf):
bar.addWidget(b)
bar.addStretch(1)
root.addLayout(bar)


self.table = QTableWidget(0,0,self)
root.addWidget(self.table, 1)
self.oos_label = QLabel("Итого брак: 0")
root.addWidget(self.oos_label)


self.binder = TableBinder(self.table)
self._presenter = presenter_cls(self)


# API для Presenter
def show_grid(self, grid, tolset, metrics):
self.binder.bind_grid(grid)
self.binder.apply_oos_coloring(grid, tolset)
self._update_metrics(metrics)


def update_tolerance_display(self, col, text, model):
# верхняя полоса не реализована в демо; просто перекрасим
self.binder.apply_oos_coloring(model.grid, model.tolset)
self._update_metrics(model.metrics)


def _update_metrics(self, m):
self.oos_label.setText(f"Итого брак: {m.bad}")


def main_table(self):
return self.table


# slots -> presenter
def _open_xlsx(self):
from .dialogs import ask_open_xlsx
p = ask_open_xlsx(self);
if p: self._presenter.open_xlsx(p)


def _open_ods(self):
from .dialogs import ask_open_ods
p = ask_open_ods(self);
if p: self._presenter.open_ods(p)


def _save_pdf(self):
from .dialogs import ask_save_pdf
p = ask_save_pdf(self, "report.pdf")
if p: self._presenter.export_pdf(p)


def _save_xlsx(self):
from .dialogs import ask_save_xlsx
p = ask_save_xlsx(self, "table.xlsx")
if p: self._presenter.save_xlsx(p)