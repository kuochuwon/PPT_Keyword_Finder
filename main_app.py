import sys

from PyQt5.QtWidgets import QApplication, QMainWindow

from configs.log_config import init_config, get_logger as logger
from ppt_finder import Ui_MainWindow


class App(QMainWindow, Ui_MainWindow):

    def __init__(self):
        super(QMainWindow, self).__init__()  # TODO 搞懂此處細節
        self.setupUi(self)
        self.show()


if __name__ == '__main__':
    init_config()
    app = QApplication(sys.argv)
    window = App()
    logger().info("Start Application...")
    try:
        sys.exit(app.exec_())
    except SystemExit:
        logger().info("Closing Application...")
