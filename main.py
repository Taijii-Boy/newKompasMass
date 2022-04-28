from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from interface import *
from kompas_api import *

import sys


class MyWin(QMainWindow):
    """Главное окно приложения"""

    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.flag = False

        # Предварительная настройка окна
        self.ui.lineEdit.setText('Закрыт')

        # Графические эффекты
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.ui.widget.setGraphicsEffect(QGraphicsDropShadowEffect(blurRadius=25, xOffset=0, yOffset=0))

        # Обработчики
        self.ui.pushButton_3.clicked.connect(lambda: self.close())
        self.ui.pushButton_4.clicked.connect(lambda: self.showMinimized())

        # События
        self.ui.label_2.mouseMoveEvent = self.move_window
        self.ui.label_3.mouseMoveEvent = self.move_window


    def move_window(self, e: QMouseEvent):
        if e.buttons() == Qt.LeftButton:
            # Перемещаем окошко
            self.move(self.pos() + e.globalPos() - self.click_position)
            self.click_position = e.globalPos()
            e.accept()


    def mousePressEvent(self, event: QMouseEvent):
        """Получаем текущие координаты курсора мыши"""
        self.click_position = event.globalPos()


    def change_status(self, new_status):
        self.ui.lineEdit.setText(new_status)


if __name__ == '__main__':

    app = QApplication(sys.argv)
    my_app = MyWin()
    my_app.show()
    kompas = KompasAPI()


    def param_changed(name, value):
        print('вошел в param_changed')
        """Что нужно делать при изменении параметра"""
        if name == 'kompas_status':
            # kompas.stop_kompas_checker()
            my_app.change_status('Открыт')

    kompas.property_changed.append(param_changed)

    # api_document = kompas.get_active_doc()
    # document = kompas.make_kompas_document(api_document)
    # my_app.ui.lineEdit_2.setText(document.name)
    # my_app.ui.lineEdit_3.setText(document.get_mass())

    sys.exit(app.exec_())
