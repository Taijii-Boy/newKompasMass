import pythoncom
import psutil
from win32com.client import Dispatch, gencache

from kompas_document import *
from cdw import *
from spw import *


class KompasAPI:

    def __init__(self):
        self.__process = self.get_proc()
        self.verify_kompas(self.__process)

    def connect_to_kompas(self):
        #  Подключим константы API Компас
        self.const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        self.const_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

        #  Подключим описание интерфейсов API5
        self.KAPI = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
        self.iKompasObject = self.KAPI.KompasObject(
            Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(self.KAPI.KompasObject.CLSID,
                                                                     pythoncom.IID_IDispatch))
        #  Подключим описание интерфейсов API7
        self.KAPI7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.application = self.KAPI7.IApplication(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.KAPI7.IApplication.CLSID,
                                                                     pythoncom.IID_IDispatch))
        self.__process = self.get_proc()


    @classmethod
    def get_proc(cls):
        for proc in psutil.process_iter():
            name = proc.name()
            if name == "KOMPAS.Exe":
                return proc

    @classmethod
    def verify_document(cls, doc):
        if not doc:
            # raise Exception("Документ Компас не найден. Откройте документ!")
            print("Документ Компас не найден. Откройте документ!")

    @classmethod
    def verify_kompas(cls, process):
        if not process:  # Компас не запущен
            # raise Exception("Вам необходимо запустить Компас!")
            print("Вам необходимо запустить Компас!")

    def close(self):
        self.__process.kill()

    def get_active_doc(self):
        """Получаем активный компас-документ"""
        document = self.application.ActiveDocument
        self.verify_document(document)
        return document

    def get_kompas_status(self) -> str:
        return "Закрыт" if not self.__process else "Открыт"

    def make_kompas_document(self, doc):
        if doc.Name.endswith('cdw'):
            return CDW(doc)
        elif doc.Name.endswith('spw'):
            return SPW(doc, self.iKompasObject)
        else:
            raise TypeError('Неверный формат. Требуется cdw или spw.')


if __name__ == '__main__':
    kompas = KompasAPI()
    kompas.connect_to_kompas()
    api_document = kompas.get_active_doc()
    document = kompas.make_kompas_document(api_document)
    print(document.extension)

    # kompas.close()
