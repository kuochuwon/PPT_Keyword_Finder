import sys
from ppt_finder import Ui_MainWindow
from PyQt5.QtWidgets import (QMainWindow,
                             QApplication)


class App(QMainWindow, Ui_MainWindow):

    def __init__(self):
        super(QMainWindow, self).__init__()  # TODO 搞懂此處細節
        self.setupUi(self)
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = App()
    # sys.exit(app.exec_())
    try:
        sys.exit(app.exec_())
    except SystemExit:
        print('Closing Window...')
