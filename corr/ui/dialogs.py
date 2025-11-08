from PyQt5.QtWidgets import QFileDialog


def ask_open_xlsx(parent):
return QFileDialog.getOpenFileName(parent, "Открыть XLSX", "", "Excel (*.xlsx)")[0]


def ask_open_ods(parent):
return QFileDialog.getOpenFileName(parent, "Открыть ODS", "", "ODS (*.ods)")[0]


def ask_save_pdf(parent, default_name: str):
return QFileDialog.getSaveFileName(parent, "Сохранить PDF", default_name, "PDF (*.pdf)")[0]


def ask_save_xlsx(parent, default_name: str):
return QFileDialog.getSaveFileName(parent, "Сохранить XLSX", default_name, "Excel (*.xlsx)")[0]