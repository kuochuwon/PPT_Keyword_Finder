# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\create_index.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

import os
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog


class Ui_Dialog(object):
    def get_powerpoint_files(self):
        # response = QFileDialog.getOpenFileNames(
        #     parent=self,
        #     caption="Select your file",
        #     directory=os.getcwd(),
        #     filter=None,
        #     initialFilter=None
        # )
        response = QFileDialog.getOpenFileNames(
            parent=QtWidgets,
            caption="Select your file",
            directory=os.getcwd()
        )
        print(response)
        # return response[0]

    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(400, 300)
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setGeometry(QtCore.QRect(40, 260, 341, 32))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(14)
        self.buttonBox.setFont(font)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(
            QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(10, 10, 221, 41))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(16)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.listWidget_2 = QtWidgets.QListWidget(Dialog)
        self.listWidget_2.setGeometry(QtCore.QRect(10, 50, 239, 78))
        self.listWidget_2.setObjectName("listWidget_2")
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(260, 60, 131, 51))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.pushButton.setFont(font)
        self.pushButton.setAutoDefault(False)
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.get_powerpoint_files)

        self.retranslateUi(Dialog)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "檔案已選取...."))
        self.pushButton.setText(_translate("Dialog", "瀏覽"))
