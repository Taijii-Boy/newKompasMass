from kompas_document import *


class SPW(KompasDocument):
    def __init__(self, doc, iKompasObject):
        super().__init__(doc)
        self.extension = 'spw'
        self.iKompasObject = iKompasObject
        self.__mass = 0.000
        self.__spc_is_revision = False
        self.__mass_increase = True

    def get_spc_interface(self) -> tuple:
        iDocumentSpc = self.iKompasObject.SpcActiveDocument()  # указатель на интерфейс документа-спецификации API5
        iSpecification = iDocumentSpc.GetSpecification()  # указатель на спецификацию
        count_rows = iDocumentSpc.ksGetSpcDocumentPagesCount() * 22  # 22 строки в листе
        return count_rows, iSpecification, iDocumentSpc

    def read_spc_spw_rows(self, count_rows) -> list:
        objects_in_spc = []

        iIter = self.iKompasObject.GetIterator()  # Получить интерфейс итератора
        iIter.ksCreateSpcIterator(self.library_path, 10, 3)
        iIter.ksMoveIterator("F")  # чтение первой строки
        iIter.ksMoveIterator("N")  # Пропускаем первые 2 строки (надпись лишняя)

        for i in range(count_rows):
            obj = iIter.ksMoveIterator("N")  # N - переход к следующему объекту
            objects_in_spc.append(obj)
        return objects_in_spc

    def get_spc_spw_date(self, objects_in_spc) -> list:
        objects = []
        for obj in objects_in_spc:
            sp_obozn = self.iSpecification.ksGetSpcObjectColumnText(obj, 4, 1, 0)
            sp_naim = self.iSpecification.ksGetSpcObjectColumnText(obj, 5, 1, 0)
            sp_col = self.iSpecification.ksGetSpcObjectColumnText(obj, 6, 1, 0)
            sp_mass = self.iSpecification.ksGetSpcObjectColumnText(obj, 8, 1, 0)
            if sp_obozn == '' and sp_naim == '':  # пропускаем пустые строки
                continue
            if sp_naim == 'Вновь устанавливаемые составные части':
                self.__mass_increase = True
                self.__spc_is_revision = True
            if sp_naim == 'Снятые составные части':
                self.__mass_increase = False
                self.__spc_is_revision = True
            objects.append((sp_naim, sp_mass, sp_col))
        return objects

    def calculate_object_mass(self, sp_naim, sp_mass, sp_col):
        if sp_naim == "Сборочный чертеж":
            pass
        elif sp_mass != '' and sp_col != '':
            if self.__mass_increase:
                self.__mass = self.__mass + float(sp_mass) * int(sp_col)
            else:
                self.__mass = self.__mass - float(sp_mass) * int(sp_col)
        elif sp_mass != '' and sp_col == '':
            if self.__mass_increase:
                self.__mass = self.__mass + float(sp_mass)
            else:
                self.__mass = self.__mass - float(sp_mass)

    def get_result(self):
        self.__mass = round(self.__mass, 3)
        return self.point_to_comma(str(self.__mass))

    def get_mass(self):
        count_rows, self.iSpecification, self.iDocumentSpc = self.get_spc_interface()
        objects_in_spc = self.read_spc_spw_rows(count_rows)
        objects = self.get_spc_spw_date(objects_in_spc)  # sp_naim, sp_mass, sp_col
        for sp_naim, sp_mass, sp_col in objects:
            self.calculate_object_mass(sp_naim, sp_mass, sp_col)
        return self.get_result()

