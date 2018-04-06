import sys
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5 import uic
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

    def updateGrouplist(self):
        cursor.execute("SELECT name,code,year FROM groups")
        grouplist = cursor.fetchall()
        print(type(grouplist))
        print(grouplist)
        for group in grouplist:
            print(type(group))
            print('-'.join(map(str, group)))
            self.ui.groupComboBox.addItem('-'.join(map(str, group)))

    def download(self):
        print("downloading...")
        html_doc = urllib.request.urlopen('https://www.mirea.ru/education/schedule-main/schedule/').read()
        soup = BeautifulSoup(html_doc, "html.parser")

        for links in soup.find_all('a'):
            if links.get('href').find("IIT-2k-17_18-vesna.xlsx") != -1:
                link = links.get('href')
                print(link)
        urllib.request.urlretrieve(link, "files/1.xlsx")

    def parse(self):
        print("parsing...")
        from openpyxl import load_workbook
        wb = load_workbook(filename='files/1.xlsx', read_only=True)
        ws = wb['Лист1']
        print(ws['F2'].value)

        for row in ws.iter_rows(min_row=4, max_row=14, min_col=6, max_col=9):
            # print(row[0].value,row[1].value,row[2].value,row[3].value)
            try:
                cursor.execute("INSERT INTO lessons VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)", (
                    'NULL', ws['F2'].value, 1, 99, 88, row[0].value, row[1].value, row[2].value, row[3].value))
                conn.commit()
            except Error as error:
                print(error)

            # for cell in row:
            #   print(cell.value)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    dbconfig = read_db_config()
    conn = MySQLConnection(**dbconfig)
    cursor = conn.cursor()

    sys.exit(app.exec_())
