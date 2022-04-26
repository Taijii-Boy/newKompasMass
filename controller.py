from model import *
from view import *
from PyQt5.QtGui import QMouseEvent

class Controller:
    def __init__(self, model, view):
        self.model = model
        self.view = view

    def start(self):
        self.view.setup(self)
        self.view.start_main_loop()

    def move_window(self, e: QMouseEvent):
        if e.buttons() == Qt.LeftButton:
            # Перемещаем окошко
            self.view.move(self.view.pos() + e.globalPos() - self.view.click_position)
            self.view.click_position = e.globalPos()
            e.accept()

    def mousePressEvent(self, event: QMouseEvent):
        """Получаем текущие координаты курсора мыши"""
        self.view.click_position = event.globalPos()



