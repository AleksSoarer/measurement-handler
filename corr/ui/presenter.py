from dataclasses import dataclass
from ..domain.types import Grid, TolSet
from ..usecases.load_table import load_from_matrix
from ..usecases.update_tolerance import update_tolerance
from ..infra.xlsx_io import load_xlsx_matrix, save_xlsx
from ..infra.ods_io import load_ods_matrix
from ..infra.pdf_print import print_table_single_page
from ..infra.pdf_merge import merge_pdfs


@dataclass
class Model:
grid: Grid
tolset: TolSet
metrics: object


class Presenter:
def __init__(self, view):
self.view = view
self.model: Model|None = None


# Загрузка
def open_xlsx(self, path: str):
matrix = load_xlsx_matrix(path)
res = load_from_matrix(matrix)
self.model = Model(res.grid, res.tolset, res.metrics)
self.view.show_grid(self.model.grid, self.model.tolset, self.model.metrics)


def open_ods(self, path: str):
matrix = load_ods_matrix(path)
res = load_from_matrix(matrix)
self.model = Model(res.grid, res.tolset, res.metrics)
self.view.show_grid(self.model.grid, self.model.tolset, self.model.metrics)


# Правка допуска
def set_tolerance(self, col: int, raw_text: str):
patch = update_tolerance(self.model.grid, self.model.tolset, col, raw_text)
self.model.metrics = patch.metrics
self.view.update_tolerance_display(col, patch.display_text, self.model)


# Экспорт
def export_pdf(self, path: str):
table = self.view.main_table()
print_table_single_page(path, table)


def save_xlsx(self, path: str):
def rgb_func(r,c):
it = self.view.table.item(r,c)
bg = it.background().color().name()[1:].upper()
fg = it.foreground().color().name()[1:].upper()
return bg, fg
save_xlsx(path, self.model.grid, rgb_func)