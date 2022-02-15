from kompas_document import *


class CDW(KompasDocument):

    def __init__(self, doc):
        super().__init__(doc)
        self.extension = 'cdw'

    def get_data_for_sp(self) -> tuple:
        sp_obozn, sp_naim, sp_col, sp_mass = [], [], [], []
        count = self.sheets_count + 1
        for sheet in range(1, count):
            iLayoutSheet = self.iLayoutSheets.ItemByNumber(sheet)
            iStamp = iLayoutSheet.Stamp

            for i in range(1014, 1225, 10):  # Обозначение
                iObozn = iStamp.Text(i)
                sp_obozn.append(iObozn.Str)

            for i in range(1015, 1226, 10):  # Наименование
                iNaim = iStamp.Text(i)
                sp_naim.append(iNaim.Str)

            for i in range(3011, 3222, 10):  # Количество
                iCol = iStamp.Text(i)
                sp_col.append(iCol.Str)

            for i in range(1019, 1230, 10):  # Масса
                iMass = iStamp.Text(i)
                sp_mass.append(iMass.Str)
        return sp_obozn, sp_naim, sp_col, sp_mass

    def get_mass_from_date(self, sp_date) -> float:
        mass = 0.000
        sp_obozn, sp_naim, sp_col, sp_mass = sp_date
        for i in range(0, len(sp_naim)):
            if sp_naim[i] == "Сборочный чертеж":
                pass
            elif sp_obozn[i] == '' and sp_naim[i] == '':  # пропускаем пустые строки
                continue
            elif sp_mass[i] != '' and sp_col[i] != '':
                mass = mass + float(self.comma_to_point(sp_mass[i])) * float(sp_col[i])
            elif sp_mass[i] != '' and sp_col[i] == '':
                mass = mass + float(self.comma_to_point(sp_mass[i]))
        return round(mass, 3)

    def get_mass(self) -> str:
        sp_date = self.get_data_for_sp()
        mass = self.get_mass_from_date(sp_date)
        return self.point_to_comma(str(mass))
