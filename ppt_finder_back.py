# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\ppt_finder.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

from create_index_2 import Ui_Form
from PyQt5 import QtCore, QtGui, QtWidgets
from ppt_functions.Powerpoint_find_keyword import PowerPoint_keyword_search

ppt_finder = PowerPoint_keyword_search()


class Ui_MainWindow(object):
    def index_window(self):
        # HINT 這邊根據create_index檔內的類別進行設定，根據情況會是QDialog or QWidget
        self.window = QtWidgets.QWidget()
        self.ui = Ui_Form()
        self.ui.setupUi(self.window)
        self.window.show()

    def search_process(self):
        text_value = self.plainTextEdit_input.toPlainText()
        print(text_value)
        result = ppt_finder.decode_find_keyword(text_value)
        self.plainTextEdit_output.insertPlainText(result)

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(883, 475)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(720, 50, 131, 51))
        # self.pushButton.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.pushButton.setFont(font)
        self.pushButton.setAutoDefault(False)
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.search_process)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 10, 221, 41))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(16)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)

        self.pushButton_2.setGeometry(QtCore.QRect(580, 50, 131, 51))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setAutoDefault(False)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.index_window)
        # self.pushButton_2.setSizePolicy(QtWidgets.QSizePolicy.Expanding,
        #                                 QtWidgets.QSizePolicy.Expanding)

        self.plainTextEdit_output = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit_output.setGeometry(QtCore.QRect(10, 170, 691, 201))
        self.plainTextEdit_output.setReadOnly(True)
        self.plainTextEdit_output.setObjectName("plainTextEdit")

        self.plainTextEdit_input = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit_input.setGeometry(QtCore.QRect(10, 50, 341, 41))
        self.plainTextEdit_input.setObjectName("textEdit")

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 883, 18))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "開始搜尋"))

        self.label.setText(_translate("MainWindow", "請輸入關鍵字(須先建立索引)"))
        self.pushButton_2.setText(_translate("MainWindow", "建立索引"))