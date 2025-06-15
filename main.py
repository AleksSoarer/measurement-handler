import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QWidget,
    QVBoxLayout, QPushButton, QLineEdit, QTableWidget, QTableWidgetItem, QLabel
)
import ezodf

class ODSViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('ODS Viewer')
        self.setGeometry(200, 200, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self.open_button = QPushButton('Открыть ODS файл')
        self.layout.addWidget(self.open_button)
        self.open_button.clicked.connect(self.open_file)

        self.range_input = QLineEdit('A1:D10')
        self.layout.addWidget(QLabel('Диапазон ячеек (например, A1:D10):'))
        self.layout.addWidget(self.range_input)

        self.show_button = QPushButton('Показать')
        self.layout.addWidget(self.show_button)
        self.show_button.clicked.connect(self.show_range)

        self.table = QTableWidget()
        self.layout.addWidget(self.table)

        self.ods_doc = None
        self.sheet = None

    def open_file(self):
        try:
            path, _ = QFileDialog.getOpenFileName(self, "Открыть ODS", "", "ODS files (*.ods)")
            print(f"Выбран файл: {path}")
            if not path:
                print("Файл не выбран.")
                return
            self.ods_doc = ezodf.opendoc(path)
            print("ODS файл успешно открыт.")
            self.sheet = self.ods_doc.sheets[0]
            print(f"Выбран лист: {self.sheet.name}, размер: {self.sheet.nrows()}x{self.sheet.ncols()}")
        except Exception as e:
            print(f"Ошибка при открытии файла: {e}")

    def show_range(self):
        if not self.sheet:
            print("Сначала откройте файл!")
            return

        rng = self.range_input.text()  # Например, "A1:D10"
        try:
            from_col, from_row, to_col, to_row = parse_range(rng)
            print(f"Диапазон: col {from_col}-{to_col}, row {from_row}-{to_row}")
            data = []
            for r in range(from_row, to_row+1):
                row_data = []
                for c in range(from_col, to_col+1):
                    try:
                        value = self.sheet[r, c].value
                    except Exception as cell_e:
                        value = None
                        print(f"Ошибка чтения ячейки {r},{c}: {cell_e}")
                    row_data.append(value)
                data.append(row_data)

            self.table.setRowCount(len(data))
            self.table.setColumnCount(len(data[0]) if data else 0)
            for i, row in enumerate(data):
                for j, cell in enumerate(row):
                    self.table.setItem(i, j, QTableWidgetItem(str(cell) if cell is not None else ""))

            print("Данные успешно выведены в таблицу.")
        except Exception as e:
            print(f"Ошибка при отображении диапазона: {e}")

def parse_range(rng):
    """ Преобразует строку типа 'A1:D10' в индексы: from_col, from_row, to_col, to_row """
    import re
    match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', rng)
    if not match:
        raise ValueError(f"Неверный формат диапазона: {rng}")
    def col2num(col):
        num = 0
        for c in col:
            num = num * 26 + (ord(c) - ord('A') + 1)
        return num - 1
    from_col, from_row, to_col, to_row = match.groups()
    return col2num(from_col), int(from_row)-1, col2num(to_col), int(to_row)-1

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = ODSViewer()
    viewer.show()
    sys.exit(app.exec_())
