from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from interface import *
from kompas_api import *

import time
import sys

# class StatusThread(QtCore.QThread):
#     """Класс потока для расчетов"""
#
#     mysignal = QtCore.pyqtSignal(str)
#
#     def __init__(self, parent=None):
#         QtCore.QThread.__init__(self, parent)
#         self.status = kompas.get_kompas_status()
#
#
#     def run(self):
#         while self.status == 'Закрыт':
#             time.sleep(1)
#             self.status = kompas.get_kompas_status()
#             kompas.connect_to_kompas()
#
#             # Передача данных из потока через сигнал
#         self.mysignal.emit('Открыт')



class MyWin(QMainWindow):
    """Главное окно приложения"""

    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.ui = Ui_Form()
        self.ui.setupUi(self)

        self.check_kompas_timer = QTimer(self)
        self.check_kompas_timer.timeout.connect(self.check_kompas)
        self.check_kompas_timer.start(3000)

        self.active_document = None
        self.kompas_status = None

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


    def check_kompas(self):
        self.kompas_status = kompas.get_kompas_status()
        my_app.ui.lineEdit.setText(self.kompas_status)
        if self.kompas_status == 'Открыт':
            kompas.connect_to_kompas()
            document = kompas.get_active_doc()
            print(document)
            if document != self.active_document:
                document = kompas.make_kompas_document(document)
                print('сделали документ')
                my_app.ui.lineEdit_2.setText(document.name)
                self.active_document = document

    def check_document(self):
        pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    my_app = MyWin()
    my_app.show()
    kompas = KompasAPI()



    # status_thread = StatusThread()
    # status_thread.mysignal.connect(my_app.add_status_text, QtCore.Qt.QueuedConnection)
    # status_thread.start()


    # api_document = kompas.get_active_doc()
    # document = kompas.make_kompas_document(api_document)
    # my_app.ui.lineEdit_2.setText(document.name)
    # my_app.ui.lineEdit_3.setText(document.get_mass())


    sys.exit(app.exec_())
