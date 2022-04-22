import sys
from PyQt5 import QtCore
from PyQt5.QtWidgets import QWidget, QApplication
from kompas_api import *

class WindowClass(QWidget):
    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.resize(300, 100)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    wc = WindowClass()
    wc.show()
    kompas = KompasAPI()
    kompas.connect_to_kompas()



    sys.exit(app.exec_())