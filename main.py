import itertools
import openpyxl
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QApplication, QTableWidgetItem, QMessageBox
from OmGTU import Ui_MainWindow
from openpyxl import load_workbook
from igraph import *
import sys
import os


file = ''
profi = []
ranks = []
wb = ws = row_count = column_count = None


def loadallpapers():
    global profi, ranks
      # Загрузка файла и считывание его данных
    # Считывание данных из таблицы в файле
    for j in range(2, row_count + 1):
        for i in range(1, column_count + 1):
            profi.append([ws.cell(row=j, column=i).value])
            ranks.append(0)
    return profi, ranks


def dict_sort(slovarik):  # Сортировка списка статей по центральности и создание списка отсортированных строк
    sorted_tuples = sorted(slovarik.items(), key=lambda item: (item[1]), reverse=True)
    sorted_dict = {k + 1: v for k, v in sorted_tuples}
    rows = list(sorted_dict.keys())
    return rows


def range_sort(slovarik):
    sorted_tuples = sorted(slovarik.items(), key=lambda item: (item[1]), reverse=True)
    sorted_dict = {k + 1: v for k, v in sorted_tuples}
    range_sorted_dict = {}
    out = {}
    i = 1
    for k, v in sorted_dict.items():
        out.setdefault(v, []).append(k)
    for k, v in out.items():
        for vv in v:
            range_sorted_dict[vv] = i
        i += 1
    return range_sorted_dict


def profile_range_sort(sorted_dict):
    range_sorted_dict = {}
    out = {}
    i = 1
    for k, v in sorted_dict.items():
        out.setdefault(v, []).append(k)
    for k, v in out.items():
        for vv in v:
            range_sorted_dict[vv] = i
        i += 1
    return range_sorted_dict


def profile_dict_sort(slovarik):  # Сортировка списка статей по центральности и создание списка отсортированных строк
    sorted_dict = {}
    sorted_keys = sorted(slovarik, key=slovarik.get)  # [1, 3, 2]

    for w in sorted_keys:
        sorted_dict[w] = slovarik[w]
    return sorted_dict


def dict_sum(slovarik0, slovarik1):
    final_slov = slovarik0.copy()
    for k, v in slovarik1.items():
        final_slov[k] = final_slov.get(k, 0) + v
    return final_slov


def degree_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.degree(g))))
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


def closeness_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.closeness(g))))
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


def betweenness_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.betweenness(g))))
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


def authority_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.authority_score(g))))
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


def hub_sort():  # Сортировка статьи по центральности и добавление её в таблицу
    g = Graph.Read_Pajek("graph.net")
    rows = dict_sort(dict(enumerate(Graph.hub_score(g))))
    data = []
    # Считывание данных из таблицы в файлк
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data


def referativ():
    g = Graph.Read_Pajek("graph.net")
    data = []
    degree = range_sort(dict(enumerate(Graph.degree(g, mode='out'))))
    closeness = range_sort(dict(enumerate(Graph.closeness(g, mode='out'))))
    hub = range_sort(dict(enumerate(Graph.hub_score(g))))
    profiles = profile_dict_sort(dict_sum(dict_sum(degree, closeness), hub))
    sorted_profiles = profile_range_sort(profiles)  # Ключ - номер статьи, значение - ранг
    rows = list(sorted_profiles.keys())
    ranks = list(sorted_profiles.values())
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j + 1, column=i).value])
    return data, ranks


def priznan():
    g = Graph.Read_Pajek("graph.net")
    data = []
    degree = range_sort(dict(enumerate(Graph.degree(g, mode='in'))))
    closeness = range_sort(dict(enumerate(Graph.closeness(g, mode='in'))))
    authority = range_sort(dict(enumerate(Graph.authority_score(g))))
    profiles = profile_dict_sort(dict_sum(dict_sum(degree, closeness), authority))
    sorted_profiles = profile_range_sort(profiles)
    rows = list(sorted_profiles.keys())
    ranks = list(sorted_profiles.values())
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j + 1, column=i).value])
    return data, ranks


def vesomost():
    g = Graph.Read_Pajek("graph.net")
    data = []
    degree = range_sort(dict(enumerate(Graph.degree(g))))
    closeness = range_sort(dict(enumerate(Graph.closeness(g))))
    betweenness = range_sort(dict(enumerate(Graph.betweenness(g))))
    authority = range_sort(dict(enumerate(Graph.authority_score(g))))
    hub = range_sort(dict(enumerate(Graph.hub_score(g))))
    profiles = profile_dict_sort(dict_sum(dict_sum(dict_sum(dict_sum(degree, closeness), betweenness), authority), hub))
    sorted_profiles = profile_range_sort(profiles)
    rows = list(sorted_profiles.keys())
    ranks = list(sorted_profiles.values())
    for j in rows:
        for i in range(1, column_count + 1):
            data.append([ws.cell(row=j+1, column=i).value])
    return data, ranks


class MyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.comboBox.addItem('Все статьи')  # Добавление названий профилей в выдвигающийся список
        self.ui.comboBox.addItem('Исходящий')
        self.ui.comboBox.addItem('Входящий')
        self.ui.comboBox.addItem('Входящий/Исходящий')
        self.ui.comboBox_2.addItem('Ключевое слово')
        self.ui.comboBox_2.addItem('Название')
        self.ui.comboBox_2.addItem('Автор')
        self.ui.comboBox_2.addItem('УДК')
        self.ui.lineEdit.setPlaceholderText("Поиск по ключевым словам...")
        self.ui.pushButton.clicked.connect(self.search)
        self.ui.comboBox.currentTextChanged.connect(self.renew)  # Фильтрация таблицы по профилю
        self.ui.action_excel.triggered.connect(self.excel_save)
        self.ui.action_2.triggered.connect(self.getfilepath)

    def getfilepath(self):
        global file, profi, ranks
        global wb, ws, row_count, column_count
        file = QtWidgets.QFileDialog.getOpenFileName()[0]
        wb = load_workbook(filename=file,
                           data_only=True)
        ws = wb.active
        row_count = ws.max_row
        column_count = ws.max_column
        profi, ranks = loadallpapers()
        self.printer(profi, ranks)

    def printer(self, mylist, ranks):  # Вывод таблицы на экран, задача кол-ва строк и столбцов, их имён
        column_names =['№ строки', 'Название', 'Авторы', 'УДК', 'Ключевые слова',
             'Издание', 'Том, выпуск, № издания', 'Год', 'Страницы', 'Ссылка']
        self.ui.tableWidget.setColumnCount(column_count)  # Задача кол-ва столбцов и строк
        self.ui.tableWidget.setHorizontalHeaderLabels(column_names)
        self.ui.tableWidget.setRowCount(row_count)
        row = 0
        col = 0
        # Заполнение таблицы в приложении
        for tup in mylist:
            for item in tup:
                self.ui.tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
                self.ui.tableWidget.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
        colPosition = 1
        self.ui.tableWidget.insertColumn(colPosition)
        column_names.insert(1, 'Ранг')
        self.ui.tableWidget.setHorizontalHeaderLabels(column_names)
        for item in ranks:
            self.ui.tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
            self.ui.tableWidget.setItem(row, colPosition, QTableWidgetItem(str(item)))
            row += 1

    def search(self):
        data = []
        row = 0
        col = 0
        rows_count = 0
        text = self.ui.lineEdit.text()  # Считывание текста введённого пользователем
        items = self.ui.tableWidget.findItems(text, QtCore.Qt.MatchContains)  # Поиск совпадений
        #  Добавление их в список
        for item in items:
            if item and item.column() == self.search_renew():
                    i = item.row()
                    for j in range(0, 11):
                        data.append(self.ui.tableWidget.item(i, j).text())
                    rows_count += 1  # запомнить количество найденных строк, для их вывода
        if data:
            #  Очистка таблицы и вывод только искомых данных
            self.ui.tableWidget.clear()
            self.ui.tableWidget.setColumnCount(11)  # Задача кол-ва столбцов и строк
            columns_headers = ['№ строки', 'Ранг', 'Название', 'Авторы', 'УДК', 'Ключевые слова',
                 'Издание', 'Том, выпуск, № издания', 'Год', 'Страницы', 'Ссылка']
            self.ui.tableWidget.setHorizontalHeaderLabels(columns_headers)
            self.ui.tableWidget.setRowCount(rows_count)
            for item in data:
                self.ui.tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
                self.ui.tableWidget.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
        else:
            self.ui.tableWidget.clear()
            self.ui.tableWidget.setColumnCount(3)  # Задача кол-ва столбцов и строк
            self.ui.tableWidget.setHorizontalHeaderLabels(
                ('Ничего', 'Не', 'Найдено')
            )
            self.ui.tableWidget.setRowCount(0)

    def search_renew(self):  # Выбор сортировки
        checker = str(self.ui.comboBox_2.currentText())
        if checker == 'Ключевое слово':
            x = 5
        if checker == 'Автор':
            x = 3
        if checker == 'УДК':
            x = 4
        if checker == 'Название':
            x = 2
        return x

    def renew(self):  # Выбор сортировки
        checker = str(self.ui.comboBox.currentText())
        if checker == 'Все статьи':
            global ranks
            self.printer(profi, ranks)
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
        if checker == 'Входящий/Исходящий':
            mylist, ranks = vesomost()
            self.printer(mylist, ranks)
        if checker == 'Исходящий':
            mylist, ranks = referativ()
            self.printer(mylist, ranks)
        if checker == 'Входящий':
            mylist, ranks = priznan()
            self.printer(mylist, ranks)

    def get_rows(self):
        columns_headers = ['№ строки', 'Ранг', 'Название', 'Авторы', 'УДК', 'Ключевые слова',
                           'Издание', 'Том, выпуск, № издания', 'Год', 'Страницы', 'Ссылка']
        rows_cnt = self.ui.tableWidget.rowCount()
        colums_cnt = self.ui.tableWidget.columnCount()
        rows = [[] for _ in range(rows_cnt)]
        for i in range(rows_cnt):
            for j in range(colums_cnt):
                rows[i].append(self.ui.tableWidget.item(i, j).text())
        rows.insert(0, columns_headers)
        return rows

    def excel_save(self):
        counter = 0
        filename = '\Ваша выборка статей'
        basename = os.environ['USERPROFILE'] + '\Desktop' + filename
        ext = 'xlsx'
        actualname = "%s.%s" % (basename, ext)
        c = itertools.count(1)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Выбранные статьи'
        rows = self.get_rows()
        check_file = os.path.exists(actualname)
        for row in rows:
            ws.append(row)
        if check_file:
            counter += 1
            actualname = "%s (%d).%s" % (basename, counter, ext)
            wb.save(actualname)
            filename = filename.replace('\\', '') + ' (' + str(counter) + ').' + ext
        else:
            wb.save(actualname)
            filename = filename.replace('\\', '')
        QMessageBox.about(self, 'Где мой файл?', 'Ваш файл на рабочем столе\n'
                          + 'Имя файла: ' + filename)



app = QApplication(sys.argv)
application = MyWindow()
application.show()
sys.exit(app.exec())