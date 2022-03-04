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
        self.timer_id = self.startTimer(3000, timerType=QtCore.Qt.VeryCoarseTimer)
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.ui.widget.setGraphicsEffect(QGraphicsDropShadowEffect(blurRadius=25, xOffset=0, yOffset=0))

        # Обработчики
        self.ui.pushButton_3.clicked.connect(lambda: self.close())
        self.ui.pushButton_4.clicked.connect(lambda: self.showMinimized())

        # Events
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

    # def add_status_text(self, status: str):
    #     self.ui.lineEdit.setText(status)

    # @staticmethod
    # def decorator_cache(some_func):
    #     mem = {}
    #     def wrapper(*args):
    #         if args in mem:
    #             print('returning from cache')
    #             return mem[args]
    #         else:
    #             result = some_func(*args)
    #             mem[args] = result
    #             return result
    #     return wrapper

    # @decorator_cache
    def _get_status_text(self):
        kompass_process = kompas.get_proc()
        return 'Открыт' if kompass_process else 'Закрыт'

    def timerEvent(self, event):
        new_status = self._get_status_text()
        my_app.ui.lineEdit.setText(new_status)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    my_app = MyWin()
    my_app.show()
    kompas = KompasAPI()
    status = kompas.get_kompas_status()
    my_app.ui.lineEdit.setText(status)

    # status_thread = StatusThread()
    # status_thread.mysignal.connect(my_app.add_status_text, QtCore.Qt.QueuedConnection)
    # status_thread.start()


    # api_document = kompas.get_active_doc()
    # document = kompas.make_kompas_document(api_document)
    # my_app.ui.lineEdit_2.setText(document.name)
    # my_app.ui.lineEdit_3.setText(document.get_mass())


    sys.exit(app.exec_())
