#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QTableWidget, QTableWidgetItem
from PyQt5 import uic
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtGui import *
from mysql.connector import MySQLConnection, Error
from bs4 import BeautifulSoup
import urllib.request
import os, os.path
import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
from configparser import ConfigParser
import re
from datetime import datetime, timedelta


def read_db_config(filename='config.ini', section='local-mysql'):  # чтение логина бд
    parser = ConfigParser()
    parser.read(filename)
    db = {}
    if parser.has_section(section):
        items = parser.items(section)
        for item in items:
            db[item[0]] = item[1]
    else:
        raise Exception('{0} not found in the {1} file'.format(section, filename))
    return db


Ui_MainWindow, QtBaseClass = uic.loadUiType("mainwindow.ui")


class MyApp(QMainWindow):
    def __init__(self):
        super(MyApp, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.dwnldButton.clicked.connect(self.download)
        self.ui.parseButton.clicked.connect(self.parse)
        self.ui.updGlButton.clicked.connect(self.update_group_list)
        self.ui.toTablesButton.clicked.connect(self.to_tables)
        self.ui.titleButton.clicked.connect(self.titles)

        self.ui.weekLabel.setText(str(datetime.now().isocalendar()[1] - 5))  # вычисление номера текущей УЧЕБНОЙ недели

        # добавление в комбобокс всех групп из базы
        dbconfig = read_db_config()
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()
        cursor.execute("SELECT name,code,year FROM groups")
        grouplist = cursor.fetchall()
        for group in grouplist:
            self.ui.groupComboBox.addItem('-'.join(map(str, group)))
        conn.close()

    def titles(self):
        self.ui.centralwidget.setCursor(QCursor(Qt.WaitCursor))
        folder = "files/all/"
        qfiles = len([name for name in os.listdir(folder) if os.path.isfile(os.path.join(folder, name))])
        print(qfiles)
        for i in range(0, qfiles):
            fpath = folder + str(i) + ".xlsx"
            print(fpath)
            wb = load_workbook(filename=fpath, read_only=True)
            for index, sheet in enumerate(wb.sheetnames):
                ws = wb[sheet]
                for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=4):
                    for cols in row:
                        value = str(cols.value)
                        if re.match(r"\bр\s*а\s*с\s*п\s*и\s*с\s*а\s*н\s*и\s*е\b", value, re.IGNORECASE):
                            print(value)

        self.ui.centralwidget.setCursor(QCursor(Qt.ArrowCursor))

    def update_group_list(self):  # получение из файла названий всех групп и запись в бд
        self.ui.centralwidget.setCursor(QCursor(Qt.WaitCursor))
        dbconfig = read_db_config()
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()

        cursor.execute("TRUNCATE TABLE groups")
        conn.commit()

        from openpyxl import load_workbook
        wb = load_workbook(filename='files/iit/IIT-2k-17_18-vesna.xlsx', read_only=True)
        ws = wb['Лист1']
        for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=200):
            for cols in row:
                if cols.value and re.match(r'\w*[-]\d\d[-]\d\d', str(cols.value)):
                    print(cols.value)
                    string = str(cols.value)
                    group = string.split("-")
                    print(string.split("-"))
                    cursor.execute("INSERT INTO groups VALUES (%s, %s, %s, %s, %s, %s, %s,%s)",
                                   (None, group[0], group[1], int(group[2]), None, None, None, None))
                    # cursor.execute("INSERT INTO groups SET name=%s", (group[0]))
                    conn.commit()

        cursor.execute("SELECT name,code,year FROM groups")
        grouplist = cursor.fetchall()
        for group in grouplist:
            self.ui.groupComboBox.addItem('-'.join(map(str, group)))

        conn.close()
        self.ui.centralwidget.setCursor(QCursor(Qt.ArrowCursor))

    def download(self):  # скачивание файла с сайта
        self.ui.centralwidget.setCursor(QCursor(Qt.WaitCursor))
        html_doc = urllib.request.urlopen('https://www.mirea.ru/education/schedule-main/schedule/').read()
        soup = BeautifulSoup(html_doc, "html.parser")
        i = 0
        for links in soup.find_all('a'):
            if links.get('href').find(".xlsx") != -1:
                link = links.get('href')
                print(link)
                print(i)
                urllib.request.urlretrieve(link, "files/all/" + str(i) + ".xlsx")
                i += 1
        """
        if not os.path.exists("files"):
            os.makedirs("files")
        if not os.path.exists("files/iit"):
            os.makedirs("files/iit")
        urllib.request.urlretrieve(link, "files/iit/IIT-2k-17_18-vesna.xlsx")"""
        self.ui.centralwidget.setCursor(QCursor(Qt.ArrowCursor))

    def to_tables(self):  # отображение данных их бд в таблицах
        dbconfig = read_db_config()
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()
        print(self.ui.groupComboBox.currentText())
        cursor.execute("SELECT type, title, teacher, room FROM lessons WHERE day=1 AND even=0 AND `group`=%s",
                       (self.ui.groupComboBox.currentText(),))
        print("exec")
        lessons = cursor.fetchall()
        conn.close()

        for i in range(6):
            for j in range(4):
                self.ui.tableWidget1.setItem(i, j, QTableWidgetItem(lessons[i][j]))
        self.ui.tableWidget1.setColumnWidth(0, 30)
        self.ui.tableWidget1.setColumnWidth(1, 170)
        self.ui.tableWidget1.setColumnWidth(2, 130)
        self.ui.tableWidget1.setColumnWidth(3, 50)

    def parse(self):  # получение из файла расписания выбранной группы и запись в бд
        self.ui.centralwidget.setCursor(QCursor(Qt.WaitCursor))
        dbconfig = read_db_config()
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()

        cursor.execute("TRUNCATE TABLE lessons")
        conn.commit()

        from openpyxl import load_workbook
        wb = load_workbook(filename='files/iit/IIT-2k-17_18-vesna.xlsx', read_only=True)
        ws = wb['Лист1']

        x = 0
        y = 0
        for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=200):
            for cols in row:
                if cols.value == self.ui.groupComboBox.currentText():
                    y = cols.row
                    x = cols.column
                    break

        mir = 4
        mar = mir + 71
        mic = x
        mac = mic + 3
        gr = ws.cell(row=y, column=x).value
        print(gr)
        number = 1
        for index, row in enumerate(ws.iter_rows(min_row=mir, max_row=mar, min_col=mic, max_col=mac)):
            title = str(row[0].value)
            subgr = 0
            day = index // 12 + 1
            if number > 6:
                number = 1
            if index % 2 == 0:
                even = 0
            else:
                even = 1
            if "(1 подгр)" in title:
                print("до: ", title)
                subgr = 1
                title = title.replace('(1 подгр)', '')
                print("после: ", title)
            if "(2 подгр)" in title:
                print("до: ", title)
                subgr = 2
                title = title.replace('(2 подгр)', '')
                print("после: ", title)
            try:
                cursor.execute("INSERT INTO lessons VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", (
                    None, gr, day, number, even, title, row[1].value, row[2].value, row[3].value,
                    None, subgr, None))
                conn.commit()
            except Error as error:
                print(error)
            number += even
        conn.close()
        self.ui.centralwidget.setCursor(QCursor(Qt.ArrowCursor))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()

    sys.exit(app.exec_())
