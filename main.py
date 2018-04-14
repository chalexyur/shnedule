import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QTableWidget, QTableWidgetItem
from PyQt5 import uic
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtGui import *
from mysql.connector import MySQLConnection, Error
from bs4 import BeautifulSoup
import urllib.request
import os
import xlrd
import xlwt
import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
from xlutils3.copy import copy
from configparser import ConfigParser
import re
from datetime import datetime, timedelta


def read_db_config(filename='config.ini', section='mysql'):
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
        self.ui.updGlButton.clicked.connect(self.updateGrouplist)
        self.ui.toTablesButton.clicked.connect(self.toTables)
        self.ui.pushButton.clicked.connect(self.Week)
        self.statusBar().addWidget(self.ui.progressBar)
        self.ui.progressBar.hide()

        self.ui.weekLabel.setText(str(datetime.now().isocalendar()[1] - 5))

        dbconfig = read_db_config()
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()
        cursor.execute("SELECT name,code,year FROM groups")
        grouplist = cursor.fetchall()
        for group in grouplist:
            self.ui.groupComboBox.addItem('-'.join(map(str, group)))
        conn.close()

    def Week(self):
        print(0)

    def updateGrouplist(self):
        from openpyxl import load_workbook
        wb = load_workbook(filename='files/1.xlsx', read_only=True)
        ws = wb['Лист1']
        for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=200):
            for cols in row:
                if cols.value and re.match(r'\w*[-]\d\d[-]\d\d', str(cols.value)):
                    print(cols.value)
                    string = str(cols.value)
                    print(string.split("-"))
        dbconfig = read_db_config()
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()
        #cursor.execute("SELECT name,code,year FROM groups")

        conn.close()

    def download(self):
        print("downloading...")
        html_doc = urllib.request.urlopen('https://www.mirea.ru/education/schedule-main/schedule/').read()
        soup = BeautifulSoup(html_doc, "html.parser")

        for links in soup.find_all('a'):
            if links.get('href').find("IIT-2k-17_18-vesna.xlsx") != -1:
                link = links.get('href')
                print(link)

        if not os.path.exists("files"):
            os.makedirs("files")
        urllib.request.urlretrieve(link, "files/1.xlsx")

    def toTables(self):
        self.ui.groupComboBox.clear()
        dbconfig = read_db_config()
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()
        cursor.execute("SELECT type, title, teacher, room FROM lessons WHERE day=1 AND even=1 AND `group`='ИКБО-06-16'")
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

    def parse(self):
        print("parsing...")
        self.ui.progressBar.show()
        self.ui.progressBar.setValue(0)
        dbconfig = read_db_config()
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()
        from openpyxl import load_workbook
        wb = load_workbook(filename='files/1.xlsx', read_only=True)
        ws = wb['Лист1']

        x = 0
        y = 0
        for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=200):
            for cols in row:
                print(cols.value)
                if (cols.value == self.ui.groupComboBox.currentText()):
                    y = cols.row
                    x = cols.column
                    print(x, y, "!!!!!!!!!!!!!!!!111111111111111111111111111111111")
                    break

        mir = 4
        mar = mir + 71
        mic = x
        mac = mic + 3
        gr = ws.cell(row=y, column=x).value
        print(gr)
        i = 0
        number = 1

        pr = 1

        for row in ws.iter_rows(min_row=mir, max_row=mar, min_col=mic, max_col=mac):
            pr += 2
            self.ui.progressBar.setValue(pr)
            day = i // 12 + 1
            if (number > 6):
                number = 1
            if (i % 2 == 0):
                even = 0
            else:
                even = 1
            try:
                cursor.execute("INSERT INTO lessons VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)", (
                    'NULL', gr, day, number, even, row[0].value, row[1].value, row[2].value, row[3].value))
                conn.commit()
            except Error as error:
                print(error)
            i += 1
            number += even
        self.ui.progressBar.setValue(100)
        conn.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()

    sys.exit(app.exec_())
