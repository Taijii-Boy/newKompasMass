import pythoncom
import psutil
from win32com.client import Dispatch, gencache


class KompasAPI:

    def __init__(self):
        self.process = self.get_proc()
        self.is_run = True if self.process else False


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
        self.process = self.get_proc()


    @staticmethod
    def get_proc():
        for proc in psutil.process_iter():
            name = proc.name()
            if name == "KOMPAS.Exe":
                return proc

    def close(self):
        if not self.is_run:
            self.process.kill()


    def get_active_document(self):
        """Получаем активный компас-документ"""
        if self.is_run:
            iDocument = self.application.ActiveDocument
            return iDocument

    def get_active_views(self):
        pass

    def get_doc_name(self, document):
        return document.Name


class KompasDocument:

    def __init__(self, doc):
        self.name = doc.Name
        self.path = doc.Path
        self.full_name = doc.PathName
        self.type = self.get_document_type(doc.DocumentType)

    def get_document_type(self, type):
        types_dict = {0: 'Неизвестный тип',
                      1: 'Чертеж',
                      2: 'Фрагмент',
                      3: 'Спецификация',
                      4: 'Деталь',
                      5: 'Сборка',
                      6: 'Текстовый документ'}
        return types_dict[type]


if __name__ == '__main__':
    kompas = KompasAPI()
    kompas.connect_to_kompas()
    document = kompas.get_active_document()
    doc = KompasDocument(document)
    print(doc.type)

    kompas.close()








