from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from interface import *
from kompas_api import *

import time
import sys


class MyWin(QMainWindow):
    """Главное окно приложения"""

    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.ui.widget.setGraphicsEffect(QGraphicsDropShadowEffect(blurRadius=25, xOffset=0, yOffset=0))
        self.timer = QTimer()
        self.timer.timeout.connect(self.add_status_text)
        self.timer.start(10000)
        self.index = 1

        # Обработчики
        self.ui.pushButton_3.clicked.connect(lambda: self.close())
        self.ui.pushButton_4.clicked.connect(lambda: self.showMinimized())

        # Events
        self.ui.label_2.mouseMoveEvent = self.move_window
        self.ui.label_3.mouseMoveEvent = self.move_window

    def move_window(self, e: QMouseEvent):
        if e.buttons() == Qt.LeftButton:
            # Move window
            self.move(self.pos() + e.globalPos() - self.click_position)
            self.click_position = e.globalPos()
            e.accept()

    def mousePressEvent(self, event: QMouseEvent):
        """Get the current position of the mouse"""
        self.click_position = event.globalPos()

    def add_status_text(self):
        self.ui.lineEdit.setText(self.index)
        self.index += 1




if __name__ == '__main__':
    app = QApplication(sys.argv)
    my_app = MyWin()
    my_app.show()



    sys.exit(app.exec_())
