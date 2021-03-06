# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\create_index_2.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

# import msvcrt
import json
import os
from pathlib import Path

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from ppt_functions.Powerpoint_find_keyword import PowerPoint_keyword_search
from configs.log_config import get_logger as logger

ppt_finder = PowerPoint_keyword_search()


class Ui_Form(object):
    def get_powerpoint_files(self):
        response = QFileDialog.getOpenFileNames(
            caption="Select your file",
            directory=os.getcwd()
        )
        print(response)
        raw_file_list = response[0]
        file_list = ppt_finder.filter(raw_file_list)

        # 在空白處加入檔名資訊
        self.listWidget_3.addItems(file_list)

    def create_keyword_library(self, Form):
        try:
            all_item = []
            count = self.listWidget_3.count()
            for i in range(count):
                all_item.append(Path(self.listWidget_3.item(i).text()))
            ppt_library = ppt_finder.convert_ppt_into_dict(all_item)
            with open("ppt_library.txt", "w") as f:
                json.dump(ppt_library, f)
                logger().info("Writing ppt_library file complete.")
                # msvcrt.getch() #會使mainwindow當掉
            QMessageBox.information(Form, "通知", "索引表已完成。")
            Form.close()
        except Exception as e:
            QMessageBox.critical(Form, "錯誤", f"索引表建置失敗。\n(您輸入的檔案之中，可能有部分檔案損毀"
                                 "，請檢查每一個檔案是否都能正常開啟。)")
            Form.close()
            logger().error(f"create_keyword_library failed: {e}")

    def message(self, Form):
        QMessageBox.information(Form, "通知", "索引表已完成。")

    def closeEvent(self, Form):
        print("QWidget closed")
        Form.close()

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 300)
        self.gridLayout = QtWidgets.QGridLayout(Form)
        self.gridLayout.setObjectName("gridLayout")
        self.label_2 = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(16)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 0, 0, 1, 1)
        spacerItem = QtWidgets.QSpacerItem(385, 17, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem, 1, 0, 1, 2)
        self.listWidget_3 = QtWidgets.QListWidget(Form)
        self.listWidget_3.setObjectName("listWidget_3")
        self.gridLayout.addWidget(self.listWidget_3, 2, 0, 1, 1)
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setAutoDefault(False)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.get_powerpoint_files)
        self.gridLayout.addWidget(self.pushButton_2, 2, 1, 1, 1)
        spacerItem1 = QtWidgets.QSpacerItem(385, 17, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem1, 3, 0, 1, 2)
        self.buttonBox = QtWidgets.QDialogButtonBox(Form)
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(14)
        self.buttonBox.setFont(font)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")

        # HINT for passing argument, using lambda expression
        self.buttonBox.accepted.connect(lambda: self.create_keyword_library(Form))
        self.buttonBox.rejected.connect(lambda: self.closeEvent(Form))
        self.gridLayout.addWidget(self.buttonBox, 4, 0, 1, 2)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.pushButton_2.setText(_translate("Form", "瀏覽"))
        self.label_2.setText(_translate("Form", "檔案已選取...."))
