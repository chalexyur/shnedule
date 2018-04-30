#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import os.path
import re
import sys
import urllib.request
from configparser import ConfigParser
from datetime import datetime

from PyQt5 import uic
from PyQt5.QtCore import Qt
from PyQt5.QtGui import *
from PyQt5.QtWidgets import QMainWindow, QApplication, QTableWidgetItem
from bs4 import BeautifulSoup
from mysql.connector import MySQLConnection, Error
from openpyxl import load_workbook
from openpyxl.compat import range


def read_db_config():
    filename = 'config.ini'
    section = 'local-mysql'
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


dbconfig = read_db_config()
conn = MySQLConnection(**dbconfig)
print(conn.is_connected())
cursor = conn.cursor()
try:
    cursor.execute("""
    create table IF NOT EXISTS groups
    (
      id          int auto_increment
        primary key,
      group_name  varchar(50) null,
      quantity    int         null,
      institute   varchar(50) null,
      last_update datetime    null,
      path_id     int         null,
      constraint groups_group_name_uindex
      unique (group_name)
    );""")
    cursor.execute("""
    create table IF NOT EXISTS lessons
    (
        id int auto_increment
            primary key,
        `group` varchar(10) null,
        day int(1) null,
        number int(1) null,
        even tinyint(1) null,
        title varchar(100) null,
        type varchar(20) null,
        teacher varchar(50) null,
        room varchar(10) null,
        weeks varchar(50) null,
        subgroup int(1) null,
        campus varchar(50) null
    );""")
    cursor.execute("""
    create table IF NOT EXISTS paths
    (
      id          int auto_increment
        primary key,
      institute   varchar(50) null,
      prog        varchar(50) null,
      course      int         null,
      ses         varchar(50) null,
      last_update datetime    null,
      past_size   int         null,
      filename    varchar(50) null,
      sheet       varchar(50) null
    );""")
    conn.commit()
except Error as error:
    print(error)


def parse_groups(worksheet, path_id):
    ws = worksheet
    for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=200):
        for cols in row:
            string = str(cols.value)
            match = re.search(r'\w*[-]\d\d[-]\d\d', string)
            if match:
                print(string)
                string = match[0]
                print(string)
                try:
                    # cursor.execute("INSERT INTO groups VALUES (%s, %s, %s, %s, %s, %s, %s,%s)",
                    # (None, group[0], group[1], int(group[2]), None, None, None, None))

                    cursor.execute("INSERT IGNORE INTO groups VALUES (%s, %s, %s, %s, %s, %s)",
                                   (None, string, None, None, None, path_id))
                    # (group[0], group[1], group[2]))
                    # cursor.execute("REPLACE INTO groups SET name=%s, code=%s, year=%s", (group[0], group[1], group[2]))
                    # cursor.execute("INSERT INTO groups SET name=%s", (group[0]))

                except Error as error:
                    print(error)
    conn.commit()


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
        self.ui.tleButton.clicked.connect(self.tle)
        self.ui.tgrButton.clicked.connect(self.tgr)
        self.ui.tpaButton.clicked.connect(self.tpa)

        self.ui.weekLabel.setText(str(datetime.now().isocalendar()[1] - 5))

        cursor.execute("SELECT group_name FROM groups")
        grouplist = cursor.fetchall()
        for group in grouplist:
            self.ui.groupComboBox.addItem('-'.join(map(str, group)))

    def titles(self):
        self.ui.centralwidget.setCursor(QCursor(Qt.WaitCursor))
        folder = "files/all/"
        qfiles = len([name for name in os.listdir(folder) if os.path.isfile(os.path.join(folder, name))])
        print(qfiles)
        for i in range(0, 5):
            fpath = folder + str(i) + ".xlsx"
            print(fpath)
            wb = load_workbook(filename=fpath, read_only=True)
            for index, sheet in enumerate(wb.sheetnames):
                ws = wb[sheet]
                for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=4):
                    for cols in row:
                        value = str(cols.value)
                        if re.match(r"\bр\s*а\s*с\s*п\s*и\s*с\s*а\s*н\s*и\s*е\b", value, re.IGNORECASE):
                            print(sheet)
                            print(value)

                            cursor.execute("INSERT INTO paths VALUES (%s,%s, %s, %s, %s, %s, %s ,%s,%s)",
                                           (None, None, None, None, None, datetime.now(), None, fpath, None))
                            conn.commit()
                            cursor.execute("SELECT LAST_INSERT_ID()")
                            path_id = cursor.fetchone()[0]
                            print(path_id)
                            parse_groups(ws, path_id)

        self.ui.centralwidget.setCursor(QCursor(Qt.ArrowCursor))

    def update_group_list(self):
        self.ui.centralwidget.setCursor(QCursor(Qt.WaitCursor))
        cursor.execute("SELECT group_name FROM groups")
        grouplist = cursor.fetchall()
        self.ui.groupComboBox.clear()
        for group in grouplist:
            self.ui.groupComboBox.addItem('-'.join(map(str, group)))
        self.ui.centralwidget.setCursor(QCursor(Qt.ArrowCursor))

    def download(self):
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

    def to_tables(self):
        print(self.ui.groupComboBox.currentText())
        cursor.execute("SELECT type, title, teacher, room FROM lessons WHERE day=1 AND even=0 AND `group`=%s",
                       (self.ui.groupComboBox.currentText(),))
        print("exec")
        lessons = cursor.fetchall()

        for i in range(6):
            for j in range(4):
                self.ui.tableWidget1.setItem(i, j, QTableWidgetItem(lessons[i][j]))
        self.ui.tableWidget1.setColumnWidth(0, 30)
        self.ui.tableWidget1.setColumnWidth(1, 170)
        self.ui.tableWidget1.setColumnWidth(2, 130)
        self.ui.tableWidget1.setColumnWidth(3, 50)

    def parse(self):
        self.ui.centralwidget.setCursor(QCursor(Qt.WaitCursor))
        from openpyxl import load_workbook
        wb = load_workbook(filename='files/all/0.xlsx', read_only=True)
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
        self.ui.centralwidget.setCursor(QCursor(Qt.ArrowCursor))

    def tle(self):
        try:
            cursor.execute("DELETE FROM paths;")
            conn.commit()
        except Error as error:
            print(error)

    def tgr(self):
        try:
            cursor.execute("DELETE FROM paths;")
            conn.commit()
        except Error as error:
            print(error)

    def tpa(self):
        try:
            cursor.execute("DELETE FROM paths;")
            conn.commit()
        except Error as error:
            print(error)

    def closeEvent(self, event):
        conn.close()
        print(conn.is_connected())
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()

    sys.exit(app.exec_())
