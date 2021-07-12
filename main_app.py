import sys
from ppt_finder import Ui_MainWindow
from create_index import Ui_Dialog
from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QWidget,
                             QPushButton,
                             QAction,
                             QLineEdit,
                             QMessageBox,
                             QPlainTextEdit)
from PyQt5 import QtCore
from PyQt5.QtCore import pyqtSlot


class App(QMainWindow, Ui_MainWindow):

    def __init__(self, parent=None):
        super(QMainWindow, self).__init__(parent)
        self.setupUi(self)
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = App()
    sys.exit(app.exec_())
