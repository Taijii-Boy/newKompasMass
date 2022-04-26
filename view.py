from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from interface import *
import sys


class View(QMainWindow):
    """Интерфейс приложения"""

    def setup(self, controller):
        QWidget.__init__(self)
        self.ui = Ui_Form()
        self.ui.setupUi(self)

        # Графические эффекты
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.ui.widget.setGraphicsEffect(QGraphicsDropShadowEffect(blurRadius=25, xOffset=0, yOffset=0))

        # Обработчики
        self.ui.pushButton_3.clicked.connect(lambda: self.close())
        self.ui.pushButton_4.clicked.connect(lambda: self.showMinimized())

        # События
        self.ui.label_2.mouseMoveEvent = controller.move_window
        self.ui.label_3.mouseMoveEvent = controller.move_window



    def start_main_loop(self):
        self.show()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    my_app = View()
    my_app.setup()
    my_app.start_main_loop()
    sys.exit(app.exec_())

