from configparser import ConfigParser
from mysql.connector import Error
from mysql.connector import MySQLConnection
from time import sleep
import urllib.request
from bs4 import BeautifulSoup
from datetime import datetime
from PyQt5.QtCore import *
import os
import os.path
import re
import sys
from time import sleep
import urllib.request
from datetime import datetime

from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from bs4 import BeautifulSoup
from mysql.connector import Error
from mysql.connector import MySQLConnection
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
    create table IF NOT EXISTS lessons
    (
         id       int auto_increment
    primary key,
  `group`  varchar(10)  null,
  day      int(1)       null,
  number   int(1)       null,
  even     tinyint(1)   null,
  title    varchar(999) null,
  type     varchar(20)  null,
  teacher  varchar(99)  null,
  room     varchar(99)  null,
  weeks    varchar(50)  null,
  subgroup int(1)       null,
  campus   varchar(99)  null
);""")
    cursor.execute("""
    create table IF NOT EXISTS paths
    (
       id          int auto_increment
    primary key,
  institute   varchar(50)  null,
  prog        varchar(50)  null,
  course      int          null,
  ses         varchar(50)  null,
  last_update datetime     null,
  past_size   int          null,
  filename    varchar(50)  null,
  sheet       varchar(50)  null,
  title       varchar(999) null,
  university  varchar(50)  null,
  groups      varchar(999) null
);""")
    cursor.execute("""
        create table IF NOT EXISTS groups
        (
          id          int auto_increment
    primary key,
  group_name  varchar(50) not null,
  quantity    int         null,
  institute   varchar(50) null,
  last_update datetime    null,
  constraint groups_group_name_uindex
  unique (group_name)
    );""")
    conn.commit()
except Error as error:
    print(error)


def parse_groups(worksheet, institute):
    ws = worksheet
    groupsstring = ""
    for row in ws.iter_rows(min_row=2, max_row=3, min_col=1, max_col=200):
        for cols in row:
            string = str(cols.value)
            match = re.search(r'\w*[-]\d\d[-]\d\d', string)
            if match:
                # print(string)
                string = match[0]
                # print(gr_code)
                # print(string)
                groupsstring += string + ','
                try:
                    # cursor.execute("INSERT INTO groups VALUES (%s, %s, %s, %s, %s, %s, %s,%s)",
                    # (None, group[0], group[1], int(group[2]), None, None, None, None))
                    cursor.execute("INSERT IGNORE INTO groups VALUES (%s, %s, %s, %s, %s)",
                                   (None, string, None, institute, None))
                    # (group[0], group[1], group[2]))
                    # cursor.execute("REPLACE INTO groups SET name=%s, code=%s, year=%s", (group[0], group[1], group[2]))
                    # cursor.execute("INSERT INTO groups SET name=%s", (group[0]))
                except Error as error:
                    print(error)
    conn.commit()
    return groupsstring


class ExecuteThread(QThread):
    my_signal = pyqtSignal()

    def run(self):
        html_page = urllib.request.urlopen('https://www.mirea.ru/education/schedule-main/schedule/').read()
        soup = BeautifulSoup(html_page, "html.parser")
        if not os.path.exists("files/"):
            os.makedirs("files/")

        # for index, link in enumerate(soup.findAll('a', attrs={'href': re.compile(".xls$")})):
        # urllib.request.urlretrieve(link.get('href'), "files/" + str(index) + ".xls")
        for index, link in enumerate(soup.findAll('a', attrs={'href': re.compile(".xlsx$")})):
            print(link.get('href'))
            urllib.request.urlretrieve(link.get('href'), "files/" + str(index) + ".xlsx")
            sleep(2)
        # for index, link in enumerate(soup.findAll('a', attrs={'href': re.compile(".pdf$")})):
        # urllib.request.urlretrieve(link.get('href'), "files/" + str(index) + ".pdf")
        self.my_signal.emit()
