from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from interface import *
from kompas_api import *

import time
import sys

class StatusThread(QtCore.QThread):
    """Поток для отслеживания состояния приложения"""

    my_signal = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)

    def run(self):
        while True:
            self.sleep(2)
            kompas_status = kompas.get_kompas_status()

            # Передача данных из потока через сигнал
            self.my_signal.emit(kompas_status)


class ActiveDocThread(QtCore.QThread):
    """Поток для отслеживания состояния приложения"""

    my_signal = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)

    def run(self):
        self.sleep(5)
        doc = kompas.get_active_doc()
        doc = kompas.make_kompas_document(doc)

        # Передача данных из потока через сигнал
        self.my_signal.emit(doc.name)


class KompasConnectThr(QtCore.QThread):

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)

    def run(self):
        kompas.connect_to_kompas()




class MyWin(QMainWindow):
    """Главное окно приложения"""

    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.ui = Ui_Form()
        self.ui.setupUi(self)

        # Запуск потоков на отслеживание состояния приложения
        # self.status_thread = StatusThread()
        # self.status_thread.my_signal.connect(self.change_status)
        # self.status_thread.start()

        # self.active_doc_thread = ActiveDocThread()
        # self.active_doc_thread.my_signal.connect(self.change_active_doc)

        self.kompas_connect_thr = KompasConnectThr()

        self.active_document = None
        self.kompas_status = None

        # Графические эффекты
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.ui.widget.setGraphicsEffect(QGraphicsDropShadowEffect(blurRadius=25, xOffset=0, yOffset=0))

        # Обработчики
        self.ui.pushButton_3.clicked.connect(lambda: self.close())
        self.ui.pushButton_4.clicked.connect(lambda: self.showMinimized())
        self.ui.lineEdit.textChanged.connect(self.what_to_do)

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
        self.kompas_status = new_status


    def change_active_doc(self, document):
        self.ui.lineEdit_2.setText(document)
        self.active_document = document


    def what_to_do(self, value):
        if value == 'Открыт':
            self.kompas_connect_thr.start()
            self.active_doc_thread.start()
            try:
                doc = kompas.get_active_doc()
                doc = kompas.make_kompas_document(doc)
                print(doc.name)
            except:
                print('че-то пошло не так')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    my_app = MyWin()
    my_app.show()
    kompas = KompasAPI()



    # kompas.connect_to_kompas()



    # status_thread = StatusThread()
    # status_thread.mysignal.connect(my_app.add_status_text, QtCore.Qt.QueuedConnection)
    # status_thread.start()


    # api_document = kompas.get_active_doc()
    # document = kompas.make_kompas_document(api_document)
    # my_app.ui.lineEdit_2.setText(document.name)
    # my_app.ui.lineEdit_3.setText(document.get_mass())


    sys.exit(app.exec_())
