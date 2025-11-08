import sys
from PyQt5.QtWidgets import QApplication
from ..ui.main_window import MainWindow
from ..ui.presenter import Presenter
from .logging_setup import setup_logging


def main():
setup_logging()
app = QApplication(sys.argv)
w = MainWindow(Presenter)
w.resize(1280, 840)
w.show()
sys.exit(app.exec_())


if __name__ == "__main__":
main()