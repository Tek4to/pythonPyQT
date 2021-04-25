from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QTableWidget, QLineEdit, QPushButton, QApplication, QMainWindow, QVBoxLayout, QWidget,\
    QTableWidgetItem
from PyQt5.QtCore import Qt
from OmGTU import Ui_MainWindow
from openpyxl import load_workbook
from igraph import *
import sys


def all_profiles():  # Загрузка всех статей
    wb = load_workbook(filename='Перечень статей.xlsx',
                       data_only=True)  # Загрузка файла и считывание его данных
    ws = wb.active
    row_count = ws.max_row
    column_count = ws.max_column
    data = []
    # Считывание данных из таблицы в файле
    for j in range(2, row_count + 1):
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j, column=i).value])
    return data


def profile():
    g = Graph.Read_Pajek("graph.net")
    slovarik = dict(enumerate(Graph.degree(g)))
    sorted_tuples = sorted(slovarik.items(), key=lambda item: (item[1]), reverse=True)
    sorted_dict = {k + 1: v for k, v in sorted_tuples}
    range_sorted_dict = sorted_dict.copy()
    i = 1
    for k in range_sorted_dict.keys():
        for k2 in range_sorted_dict.values():
            if k != k2 and range_sorted_dict.get(k2) == range_sorted_dict.get(k2):
                range_sorted_dict[k] = i
            else:
                range_sorted_dict[k] = i
                i += 1


profile()


def dict_sort(slovarik):  # Сортировка списка статей по центральности и создание списка отсортированных строк
    sorted_tuples = sorted(slovarik.items(), key=lambda item: (item[1]), reverse=True)
    sorted_dict = {k + 1: v for k, v in sorted_tuples}
    rows = list(sorted_dict.keys())
    return rows


def degree_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.degree(g))))
    wb = load_workbook(filename='Перечень статей.xlsx',
                       data_only=True)  # Загрузка файла и считывание его данных
    ws = wb.active
    column_count = ws.max_column
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


def closeness_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.closeness(g))))
    wb = load_workbook(filename='Перечень статей.xlsx',
                       data_only=True)  # Загрузка файла и считывание его данных
    ws = wb.active
    column_count = ws.max_column
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


def betweenness_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.betweenness(g))))
    wb = load_workbook(filename='Перечень статей.xlsx',
                       data_only=True)  # Загрузка файла и считывание его данных
    ws = wb.active
    column_count = ws.max_column
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


def authority_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.authority_score(g))))
    wb = load_workbook(filename='Перечень статей.xlsx',
                       data_only=True)  # Загрузка файла и считывание его данных
    ws = wb.active
    column_count = ws.max_column
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


def hub_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.hub_score(g))))
    wb = load_workbook(filename='Перечень статей.xlsx',
                       data_only=True)  # Загрузка файла и считывание его данных
    ws = wb.active
    column_count = ws.max_column
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


class MyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.comboBox.addItem('Все профили')  # Добавление названий профилей в выдвигающийся список
        self.ui.comboBox.addItem('По степени связанности')
        self.ui.comboBox.addItem('По близости')
        self.ui.comboBox.addItem('По посредничеству')
        self.ui.comboBox.addItem('По авторитетности')
        self.ui.comboBox.addItem('По концентрации')
        self.ui.comboBox.addItem('Реферативность')
        self.ui.comboBox.addItem('Признанность')
        self.ui.comboBox.addItem('Весомость')
        self.ui.lineEdit.setPlaceholderText("Поиск по ключевым словам...")
        self.ui.pushButton.clicked.connect(self.search)
        self.ui.pushButton_3.clicked.connect(self.renew)  # Фильтрация таблицы по профилю
        self.printer(all_profiles())

    def printer(self, mylist):  # Вывод таблицы на экран, задача кол-ва строк и столбцов, их имён
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

    def search(self):
        s = self.ui.lineEdit.text()
        data = []
        # Считывание данных из таблицы в файле
        for j in range(self.ui.tableWidget.rowCount()):
            for i in range(self.ui.tableWidget.columnCount()):
                data.append(self.ui.tableWidget.item(j, i).text())
        print(data)

    def renew(self):  # Выбор сортировки
        checker = str(self.ui.comboBox.currentText())
        if checker == 'Все профили':
            self.printer(all_profiles())
        if checker == 'По степени связанности':
            self.printer(degree_sort())
        if checker == 'По близости':
            self.printer(closeness_sort())
        if checker == 'По посредничеству':
            self.printer(betweenness_sort())
        if checker == 'По авторитетности':
            self.printer(authority_sort())
        if checker == 'По концентрации':
            self.printer(hub_sort())


app = QApplication(sys.argv)
application = MyWindow()
application.show()
sys.exit(app.exec())