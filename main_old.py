# -*- coding: utf-8 -*-
# |1

import pythoncom, pyperclip
from win32com.client import Dispatch, gencache
import tkinterWind as tw
from tkinter import messagebox
import os


class Main_logic():

    def __init__(self):
        self.app = tw.Application()  # Подключаем UI
        self.status_open_document = self.get_Kompas()  # Пробуем подключиться к компасу
        self.change_status_in_app()  # Отображаем статус компаса в приложении
        self.app.btn_get_mass.configure(command=self.get_mass)

    def close_Kompas_process(self, program="KOMPAS.exe"):  # Принудительно убиваем процесс после закрытия
        os.system("TASKKILL /F /IM " + program)

    def get_Kompas(self):
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

        #  Получим активный документ
        self.iDocument = self.application.ActiveDocument

        if self.iDocument:
            return True  # Компас запущен
        else:
            self.close_Kompas_process()
            return False  # Компас не запущен

    def change_status_in_app(self):
        if self.status_open_document:
            self.app.label_status.configure(text="Открыт", fg="green")
        else:
            self.app.label_status.configure(text="Закрыт", fg="red")

    def get_mass(self):
        self.app.save_configuration()  # Сохраняем данные Path в файл config для открытия в будущем
        if self.app.combo_document_type.get() == "*.cdw":
            self.get_mass_from_cdw()
        elif self.app.combo_document_type.get() == "*.spw":
            self.get_mass_from_spw()
        elif self.app.combo_document_type.get() == "Служебка":
            self.get_mass_from_table()

    # Функция, меняющая запятую в значениях на точку
    def point(self, text):
        text = text.replace(',', '.')
        return text

    # Функция, меняющая точку в значениях на запятую
    def comma(self, text):
        text = text.replace('.', ',')
        return text

    #############################################
    #                     CDW                     #
    #############################################

    def get_mass_from_cdw(self):

        specification_date = ()

        # Функция получения данных из колонок спецификации
        def get_data_for_cdw():
            sp_obozn = []
            sp_naim = []
            sp_col = []
            sp_mass = []
            sp_date = ()

            count = iLayoutSheets.Count + 1  # Количество листов + 1 для цикла

            for sheet_num in range(1, count):  # Перебираем листы
                iLayoutSheet = iLayoutSheets.ItemByNumber(sheet_num)
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

            sp_date = (sp_obozn, sp_naim, sp_col, sp_mass)
            return sp_date

        # Функция подсчета массы
        def get_mass_cdw(sp_date):
            mass = 0.000
            sp_obozn, sp_naim, sp_col, sp_mass = sp_date

            for i in range(0, len(sp_naim)):
                if sp_naim[i] == "Сборочный чертеж":
                    pass
                elif sp_obozn[i] == '' and sp_naim[i] == '':  # пропускаем пустые строки
                    continue
                elif sp_mass[i] != '' and sp_col[i] != '':
                    mass = mass + float(self.point(sp_mass[i])) * float(sp_col[i])
                elif sp_mass[i] != '' and sp_col[i] == '':
                    mass = mass + float(self.point(sp_mass[i]))
            return mass

        iLayoutSheets = self.iDocument.LayoutSheets
        specification_date = get_data_for_cdw()  # Считываем данные спецификации
        mass = round(get_mass_cdw(specification_date), 3)  # Считаем массу и округляем до 4 знаков после запятой
        if mass != 0:
            mass = self.comma(str(mass))  # Меняем точку на запятую для вывода
            self.app.label_mass_result.configure(text=f"{mass} кг", font=("Arial", 10, "bold"))  # Выводим в приложение
            pyperclip.copy(mass)  # Копируем в буфер обмена
        else:
            messagebox.showwarning("Внимание", "Спецификация .cdw не найдена")

            #############################################

    #                   SPW                     #
    #############################################

    def get_mass_from_spw(self):
        mass = 0.000
        spc_is_revision = False
        mass_increase = True
        self.app.label_mass.configure(text="Общая масса равна: ")  # Возвращаем исходное состояние

        # Подсчет массы спецификации
        def get_mass(sp_naim, sp_mass, sp_col):
            nonlocal mass
            if sp_naim == "Сборочный чертеж":
                pass
            elif sp_mass != '' and sp_col != '':
                if mass_increase == True:
                    mass = mass + float(sp_mass) * float(sp_col)
                else:
                    mass = mass - float(sp_mass) * float(sp_col)
            elif sp_mass != '' and sp_col == '':
                if mass_increase == True:
                    mass = mass + float(sp_mass)
                else:
                    mass = mass - float(sp_mass)

        # Функция вывода результата на экран и в буфер обмена
        def get_result(mass, spc_is_revision):
            mass = round(mass, 3)  # Округляем с точностью до 3 знаков после запятой
            if spc_is_revision:
                self.app.label_mass.configure(text="Изменение массы: ")
                if mass > 0:
                    mass = self.comma(str(mass))  # Меняем точку на запятую для вывода
                    self.app.label_mass_result.configure(text=f"+{mass} кг",
                                                         font=("Arial", 10, "bold"))  # Выводим в приложение
                    pyperclip.copy(f"+{mass}")  # Копируем в буфер обмена
                else:
                    mass = self.comma(str(mass))  # Меняем точку на запятую для вывода
                    self.app.label_mass_result.configure(text=f"{mass} кг",
                                                         font=("Arial", 10, "bold"))  # Выводим в приложение
                    pyperclip.copy(mass)  # Копируем в буфер обмена
            else:
                mass = self.comma(str(mass))  # Меняем точку на запятую для вывода
                self.app.label_mass_result.configure(text=f"{mass} кг",
                                                     font=("Arial", 10, "bold"))  # Выводим в приложение
                pyperclip.copy(mass)  # Копируем в буфер обмена

        def get_spc_spw_date(objects_in_spc):
            nonlocal spc_is_revision
            for obj in objects_in_spc:
                sp_obozn = iSpecification.ksGetSpcObjectColumnText(obj, 4, 1, 0)
                sp_naim = iSpecification.ksGetSpcObjectColumnText(obj, 5, 1, 0)
                sp_col = iSpecification.ksGetSpcObjectColumnText(obj, 6, 1, 0)
                sp_mass = iSpecification.ksGetSpcObjectColumnText(obj, 8, 1, 0)
                if sp_obozn == '' and sp_naim == '':  # пропускаем пустые строки
                    continue
                if sp_naim == 'Вновь устанавливаемые составные части':
                    mass_increase = True
                    spc_is_revision = True
                if sp_naim == 'Снятые составные части':
                    mass_increase = False
                    spc_is_revision = True
                get_mass(sp_naim, sp_mass, sp_col)
            get_result(mass, spc_is_revision)

        def del_spc_spw_objects(objects_in_spc, iSpecification, iDocumentSpc):
            objects_to_del = []
            # indexes = []
            # # indexes_list = []
            # object_list = ["Сборочные единицы", "Детали", "Стандартные изделия", "Болты",
            #                 "Винты", "Втулки"]
            spc_spw_list = []

            for obj in enumerate(objects_in_spc):
                sp_col = iSpecification.ksGetSpcObjectColumnText(obj[1], 6, 1, 0)
                sp_naim = iSpecification.ksGetSpcObjectColumnText(obj[1], 5, 1, 0)
                # if sp_naim in object_list:
                #     indexes.append(obj[0])
                spc_spw_list.append([obj[0], obj[1], sp_naim, sp_col])

            objects_to_del = [x for x in spc_spw_list if x[3] in self.app.load_configuration()[1]]  # [1111, 2222]
            for obj in objects_to_del:
                iDocumentSpc.ksDeleteObj(obj[1])

            # TODO
            # for i in range(len(indexes)):
            #     if indexes[i] != indexes[len(indexes)-1]:
            #         indexes_list.append((indexes[i], indexes[i+1]))
            # print(indexes_list)

            # for i in range(len(indexes_list)):
            #     for j in range(indexes_list[i][0], indexes_list[i][1]+1):
            #         print(j)
            #         if spc_list[j][0] == j:
            #             print(spc_list[j])
            #     print("--------")

        def get_spc_interface():
            try:
                iSpecificationDocument = self.KAPI7.ISpecificationDocument(
                    self.iDocument)  # интерфейс документа-спецификации API7
                iDocumentSpc = self.iKompasObject.SpcActiveDocument()  # указатель на интерфейс документа-спецификации API5
                iSpecification = iDocumentSpc.GetSpecification()  # указатель на спецификацию
                count_rows = iDocumentSpc.ksGetSpcDocumentPagesCount() * 22  # Узнаем количество листов спец-ции и умножаем на 22 строки в листе
                return (count_rows, iSpecification, iDocumentSpc)
            except:
                messagebox.showwarning("Внимание", "Спецификация не найдена")
                return

        def read_spc_spw_rows(count_rows, iSpecification, iDocumentSpc):
            objects_in_spc = []

            path = self.app.entry_library_path.get()  # Путь к библиотеке стилей
            iIter = self.iKompasObject.GetIterator()  # Получить интерфейс итератора
            iIter.ksCreateSpcIterator(path, 10, 3)  # Создать итератор по объектам спецификации
            obj = iIter.ksMoveIterator("F")  # чтение первой строки
            obj = iIter.ksMoveIterator(
                "N")  # В нашей спецификации пропускаем первые 2 строки (там че-та какая-та надпись лишняя)

            # Цикл для чтения строки
            for i in range(count_rows):
                obj = iIter.ksMoveIterator("N")  # N - переход к следующему объекту
                objects_in_spc.append(obj)
            return objects_in_spc

        # MAIN
        count_rows, iSpecification, iDocumentSpc = get_spc_interface()
        objects_in_spc = read_spc_spw_rows(count_rows, iSpecification, iDocumentSpc)

        if self.app.need_to_delete_value.get() == 1:
            del_spc_spw_objects(objects_in_spc, iSpecification, iDocumentSpc)

        get_spc_spw_date(objects_in_spc)

    #############################################
    #                TABLE - Таблица           #
    #############################################

    def get_mass_from_table(self):
        total_mass = 0.000

        # Функция считывает таблицу построчно, возвращает списки строк в rows_in_table
        def get_rows_from_table():
            rows_in_table = []
            current_row = []
            for row in range(2, iTable.RowsCount):
                for column in range(3, 8):
                    iTableCell = iTable.Cell(row, column)
                    iTextObject = iTableCell.Text
                    iText = iTextObject._oleobj_.QueryInterface(self.KAPI7.IText.CLSID, pythoncom.IID_IDispatch)
                    iText = self.KAPI7.IText(iText)
                    current_row.append(iText.Str)
                rows_in_table.append(current_row)
                current_row = []
            return rows_in_table

        # Функция удаляет пустые строки
        def del_empty_rows(table_list):
            emty_template = ["", "", "", "", ""]
            for row in table_list:
                if row == emty_template:
                    table_list.remove(row)
            return table_list

        # Функция подсчета массы
        # 0 - Обозначение, 1 - Наименование, 2 - Кол на сборке, 3 - Количество на изд, 4 - Масса
        def get_table_mass(table_list):
            mass = 0.000
            mass_increase = True  # Флаг, показывающий, суммировать массу или вычитать
            for row in table_list:
                if row[1] == "Изменение массы":
                    continue
                elif row[1] == "Снять" or row[1] == "снять":
                    mass_increase = False  # Массу будем вычитать
                elif row[1] == "Добавить" or row[1] == "добавить":
                    mass_increase = True  # Массу будем складывать
                elif row[4] != "" and row[3] != "":
                    if mass_increase:
                        mass = mass + float(self.point(row[4])) * float(row[3])
                    else:
                        mass = mass - float(self.point(row[4])) * float(row[3])
                elif row[4] != "" and row[3] == "":
                    if mass_increase:
                        mass = mass + float(self.point(row[4]))
                    else:
                        mass = mass - float(self.point(row[4]))
            return mass

        # Функция вывода результата на экран и в буфер обмена
        def get_result_from_table(mass):
            self.app.label_mass.configure(text="Изменение массы: ")
            if mass > 0:
                mass = self.comma(str(mass))  # Меняем точку на запятую для вывода
                # self.iKompasObject.ksMessage("Изменение массы: +" + mass + " кг") # Выдаем сообщение в компасе
                self.app.label_mass_result.configure(text=f"+{mass} кг",
                                                     font=("Arial", 10, "bold"))  # Выводим в приложение
                pyperclip.copy(f"+{mass}")  # Копируем в буфер обмена
            else:
                mass = self.comma(str(mass))  # Меняем точку на запятую для вывода
                # self.iKompasObject.ksMessage("Изменение массы: " + mass + " кг") # Выдаем сообщение в компасе
                self.app.label_mass_result.configure(text=f"{mass} кг",
                                                     font=("Arial", 10, "bold"))  # Выводим в приложение
                pyperclip.copy(mass)  # Копируем в буфер обмена

        # Удаляем из таблицы лишние строки
        def clear_table():
            rows_to_delete = []
            iDrawingObject = self.KAPI7.IDrawingObject(iTable)
            for row in range(2, iTable.RowsCount):  # Проходим по всем строкам, не включая шапку
                for column in (4, 6, 8):
                    if column == 6:
                        iTableCell = iTable.Cell(row, column)  # 6 - Номер колонки с количеством на изделии
                        iTextObject = iTableCell.Text
                        iText = iTextObject._oleobj_.QueryInterface(self.KAPI7.IText.CLSID, pythoncom.IID_IDispatch)
                        iText = self.KAPI7.IText(iText)
                        iText = iText.Str
                        if iText in self.app.load_configuration()[1]:  # 1111, 2222
                            rows_to_delete.append(iTableCell.Row)  # Создаем список строк, которые необходимо удалить
            first_iter = True
            j = 1
            for r in rows_to_delete:
                if first_iter:
                    iTable.DeleteRow(r)  # Удаляем строки
                    first_iter = False
                else:
                    iTable.DeleteRow(r - j)  # Учитываем удаленные строки в списке rows_to_delete
                    j += 1
            iDrawingObject.Update()  # Обновляем графический объект

        try:
            iKompasDocument2D = self.KAPI7.IKompasDocument2D(self.iDocument)
            iDocument2D = self.iKompasObject.ActiveDocument2D()
            iKompasDocument2D1 = self.KAPI7.IKompasDocument2D1(iKompasDocument2D)
        except:
            messagebox.showwarning("Внимание", "Служебная записка не найдена")
            return

        # Получаем выделенную таблицу

        iSelectionManager = iKompasDocument2D1.SelectionManager
        selected_object = iSelectionManager.SelectedObjects

        if type(selected_object) != tuple and (
                selected_object == None or selected_object.Type != 13062):  # Если ничего не выделили или выделена не таблица
            messagebox.showwarning("Внимание", "Вы должны выделить одну или несколько таблиц")
        elif type(selected_object) == tuple:
            for sel_object in selected_object:
                if sel_object.Type == 13062:  # Если в выделенных элементах присутствует Таблица
                    iTable = self.KAPI7.ITable(sel_object)
                    rows_in_table = get_rows_from_table()  # Получаем списком данные из таблицы
                    if self.app.need_to_delete_value.get() == 1:  # Если стоит галочка Удалять лишние строки
                        clear_table()  # Удаляем лишние строки
                        rows_in_table = get_rows_from_table()  # Повторно запрашиваем данные очищенной таблицы
                        rows_in_table = del_empty_rows(rows_in_table)  # Удаляем лишние строки
                    mass = round(get_table_mass(rows_in_table),
                                 3)  # Считаем массу и округляем до 3 знаков после запятой
                    total_mass = total_mass + mass
                else:
                    continue  # Пропускаем все выделенные объекты кроме таблиц
            total_mass = round(total_mass, 3)  # Округляем до 3 знаков после запятой
            get_result_from_table(total_mass)  # Выводим результат
        else:
            # Если выделена одна таблица
            iTable = self.KAPI7.ITable(selected_object)
            rows_in_table = get_rows_from_table()  # Получаем списком данные из таблицы
            if self.app.need_to_delete_value.get() == 1:  # Если стоит галочка Удалять лишние строки
                clear_table()  # Удаляем лишние строки
                rows_in_table = get_rows_from_table()  # Повторно запрашиваем данные очищенной таблицы
                rows_in_table = del_empty_rows(rows_in_table)  # Удаляем лишние строки
            mass = round(get_table_mass(rows_in_table), 3)  # Считаем массу и округляем до 3 знаков после запятой
            get_result_from_table(mass)  # Выводим результат


if __name__ == '__main__':
    application_logic = Main_logic()
    application_logic.app.attributes("-topmost", True)  # Запускать поверх всех окон
    application_logic.app.mainloop()