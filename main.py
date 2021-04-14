from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QTableWidgetItem
from OmGTU import Ui_MainWindow
from openpyxl import load_workbook
import sys


def all_profiles():
    wb = load_workbook(filename='Перечень статей.xlsx',
                       data_only=True)  # Загрузка файла и считывание его данных
    ws = wb.active
    row_count = ws.max_row
    column_count = ws.max_column
    data = []
    # Считывание данных из таблицы в файлк
    for j in range(2, row_count + 1):
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j, column=i).value])
    return data


class MyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.comboBox.addItem("Все профили")  # Добавление названий профилей в выдвигающийся список
        self.ui.comboBox.addItem("Реферативность")
        self.ui.comboBox.addItem("Признанность")
        self.ui.comboBox.addItem("Весомость")

        self.ui.label1.setFont(
            QtGui.QFont('TimesNewRoman', 14))  # Шрифт и его размер в названии

        self.ui.pushButton_3.clicked.connect(self.renew)  # Фильтрация таблицы по профилю

    def printer(self, mylist):
        wb = load_workbook(filename='Перечень статей.xlsx', data_only=True)  # Загрузка файла и считывание его данных
        ws = wb.active
        row_count = ws.max_row
        column_count = ws.max_column
        self.ui.tableWidget.setColumnCount(column_count)  # Задача кол-ва столбцов и строк
        self.ui.tableWidget.setHorizontalHeaderLabels(
            ('№ статьи', 'Название', 'Авторы', 'УДК', 'Ключевые слова',
             'Издание', 'Том, выпуск, № издания', 'Год', 'Страницы', 'Ссылка')
        )
        self.ui.tableWidget.setRowCount(row_count)

        row = 0
        col = 0

        # Заполнение таблицы в приложении
        for tup in mylist:
            for item in tup:
                self.ui.tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
                self.ui.tableWidget.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1

    def renew(self):
        checker = str(self.ui.comboBox.currentText())
        if checker == 'Все профили':
            self.printer(all_profiles())

        if checker == 'Реферативность':
            self.printer()

        # if checker == 'Признанность':
        # if checker == 'Весомость':


app = QtWidgets.QApplication([])
application = MyWindow()
application.show()
sys.exit(app.exec())