import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import QObject, pyqtSignal
import sqlite3
import random
import pyperclip
import keyboard

form_class = uic.loadUiType("inputer.ui")[0]

class KeyboardManager(QObject):
    pasteSignal = pyqtSignal()

    def start(self):
        keyboard.add_hotkey("ctrl+v", self._ctrl_v_callback)

    def _ctrl_v_callback(self):
        self.pasteSignal.emit()

class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.pb1.clicked.connect(self.bt1)
        self.pb2.clicked.connect(self.bt2)
        self.pb3.clicked.connect(self.bt3)
        self.pb4.clicked.connect(self.bt4)
        self.pb5.clicked.connect(self.bt5)
        self.con = sqlite3.connect("inputer.db")
        self.cursor = self.con.cursor()
        keyboard_manager = KeyboardManager(self)
        keyboard_manager.pasteSignal.connect(self.kt)
        keyboard_manager.start()

    def bt1(self):
        self.lb1.clear()
        self.tb.clear()
        self.lb1.setText("공통")
        self.cursor.execute("SELECT max(rowid) FROM common")
        maxrow = self.cursor.fetchone()
        r = random.randrange(1, maxrow[0] + 1)
        self.cursor.execute("SELECT * FROM common WHERE rowid =?", (r,))
        t = self.cursor.fetchone()
        self.tb.setPlainText(t[0])
        pyperclip.copy(t[0] + "\n")

    def bt2(self):
        self.lb1.clear()
        self.tb.clear()
        self.lb1.setText("SPR")
        self.cursor.execute("SELECT max(rowid) FROM spr")
        maxrow = self.cursor.fetchone()
        r = random.randrange(1, maxrow[0] + 1)
        self.cursor.execute("SELECT * FROM spr WHERE rowid =?", (r,))
        t = self.cursor.fetchone()
        self.tb.setPlainText(t[0])
        pyperclip.copy(t[0] + "\n")

    def bt3(self):
        self.lb1.clear()
        self.tb.clear()
        self.lb1.setText("OBS")
        self.cursor.execute("SELECT max(rowid) FROM obs")
        maxrow = self.cursor.fetchone()
        r = random.randrange(1, maxrow[0] + 1)
        self.cursor.execute("SELECT * FROM obs WHERE rowid =?", (r,))
        t = self.cursor.fetchone()
        self.tb.setPlainText(t[0])
        pyperclip.copy(t[0] + "\n")

    def bt4(self):
        self.lb1.clear()
        self.tb.clear()
        self.lb1.setText("AD")
        self.cursor.execute("SELECT max(rowid) FROM ad")
        maxrow = self.cursor.fetchone()
        r = random.randrange(1, maxrow[0] + 1)
        self.cursor.execute("SELECT * FROM ad WHERE rowid =?", (r,))
        t = self.cursor.fetchone()
        self.tb.setPlainText(t[0])
        pyperclip.copy(t[0] + "\n")

    def bt5(self):
        self.lb1.clear()
        self.tb.clear()
        self.lb1.setText("Dep")
        self.cursor.execute("SELECT max(rowid) FROM dep")
        maxrow = self.cursor.fetchone()
        r = random.randrange(1, maxrow[0] + 1)
        self.cursor.execute("SELECT * FROM dep WHERE rowid =?", (r,))
        t = self.cursor.fetchone()
        self.tb.setPlainText(t[0])
        pyperclip.copy(t[0] + "\n")

    def kt(self):
        if self.lb1.text() == "공통":
            self.bt1()
        elif self.lb1.text() == "SPR":
            self.bt2()
        elif self.lb1.text() == "OBS":
            self.bt3()
        elif self.lb1.text() == "AD":
            self.bt4()
        elif self.lb1.text() == "Dep":
            self.bt5()

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()