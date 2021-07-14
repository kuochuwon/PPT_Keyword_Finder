# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\create_index_2.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

import os
import msvcrt
import json
from pathlib import Path
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from ppt_functions.Powerpoint_find_keyword import PowerPoint_keyword_search

ppt_finder = PowerPoint_keyword_search()


class Ui_Form(object):
    def get_powerpoint_files(self):
        # ppt_finder = PowerPoint_keyword_search()
        response = QFileDialog.getOpenFileNames(
            caption="Select your file",
            directory=os.getcwd()
        )
        print(response)
        raw_file_list = response[0]
        file_list = ppt_finder.create_filelist(raw_file_list)

        # 在空白處加入檔名資訊
        self.listWidget_3.addItems(file_list)
        # return file_list  # TODO Wrong

    def create_keyword_library(self, Form):
        all_item = []
        count = self.listWidget_3.count()
        for i in range(count):
            all_item.append(Path(self.listWidget_3.item(i).text()))
        ppt_library = ppt_finder.extractwords_into_dict(all_item)
        with open("ppt_library.txt", "w") as f:
            json.dump(ppt_library, f)
            print("write file complete")
            # msvcrt.getch() #會使mainwindow當掉
        Form.close()

    # TODO error happened
    def closeEvent(self, Form):
        print("QWidget closed")
        Form.close()

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 300)
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(290, 90, 71, 31))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setAutoDefault(False)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(
            self.get_powerpoint_files)  # TODO Wrong
        self.buttonBox = QtWidgets.QDialogButtonBox(Form)
        self.buttonBox.setGeometry(QtCore.QRect(10, 210, 341, 32))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(14)
        self.buttonBox.setFont(font)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(
            QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")

        self.buttonBox.accepted.connect(lambda: self.create_keyword_library(Form))

        # HINT for passing argument, using lambda expression
        self.buttonBox.rejected.connect(
            lambda: self.closeEvent(Form))

        self.listWidget_3 = QtWidgets.QListWidget(Form)  # HINT 文字輸入方塊
        self.listWidget_3.setGeometry(QtCore.QRect(30, 70, 239, 78))
        self.listWidget_3.setObjectName("listWidget_3")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(30, 20, 221, 41))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(16)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.pushButton_2.setText(_translate("Form", "瀏覽"))
        self.label_2.setText(_translate("Form", "檔案已選取...."))
