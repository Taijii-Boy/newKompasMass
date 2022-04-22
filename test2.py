# Для глобальной переменной это невозможно с использованием назначения.
# Но для атрибута это довольно просто: просто используйте свойство (или, может быть, __getattr__/__setattr__).
# Это фактически превращает назначение в вызов функции, позволяя вам добавить любое дополнительное поведение,
# которое вам нравится:


class Foo(QWidget):
    valueChanged = pyqtSignal(object)

    def __init__(self, parent=None):
        super(Foo, self).__init__(parent)
        self._t = 0

    @property
    def t(self):
        return self._t

    @t.setter
    def t(self, value):
        self._t = value
        self.valueChanged.emit(value)

# Теперь вы можете сделать foo = Foo(); foo.t = 5 foo = Foo(); foo.t = 5 и сигнал будет испускаться после
# изменения значения.
