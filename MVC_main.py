import sys
from PyQt5.QtWidgets import QApplication
from model import *
from view import *
from controller import *


def main():
    app = QApplication(sys.argv)
    # создаём модель
    model = Model()
    # создаем представление
    view = View()
    # создаём контроллер и передаём ему ссылку на модель
    controller = Controller(model, view)
    controller.start()
    app.exec_()


if __name__ == '__main__':
    sys.exit(main())
