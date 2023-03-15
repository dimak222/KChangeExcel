# -*- coding: cp1251 -*-
# Групповое заполнение свойств 3d моделей из имён их файлов

import pythoncom, os, openpyxl
from win32com.client import Dispatch, gencache
try:
    # импорт модуля для python 2.xx
    from tkFileDialog import *
except:
    # импорт модуля для python 3.xx
    from tkinter.filedialog import *

kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
kompas_api_object = kompas_api7_module.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
iApplication = kompas_api_object.Application

# Запрос указания папки
directory = askdirectory()
if directory:
    # Получаем список всех файлов имеющихся в папке
    all_files = os.listdir(directory)
    # Фильтруем файлы, отбирая только 3D-модели
    files = filter(lambda x: x.endswith('.m3d'), all_files)

wb = openpyxl.reader.excel.load_workbook(filename="Изменение обозначения и наименования v0.1.xlsx", data_only=True)
wb.active = 0
sheet = wb.active
print(sheet['С18'].value)

    if files:
        # Создаём цикл для работы с каждой 3D-моделью
        for f in files:
            # Формируем полный путь к файлу
            PathName = directory + '/' + f
            iDocuments = iApplication.Documents
            print(PathName)

            # Открываем файл КОМПАСом
##            iKompasDocument = iDocuments.Open (PathName, False, False )# False - в невидимом режиме, False - с возможностью редактирования
##            iKompasDocument3D = kompas_api7_module.IKompasDocument3D(iKompasDocument)
##
##            # В невидимом режиме iKompasDocument.Name возвращает полный путь, поэтому обрезаем лишнее
##            name = iKompasDocument.Name.split('/')[-1]
##            # срез имени без последних 4-x символов
##            name = name[:-4]
##            # разбиваем на список по пробелам
##            name = name.split('_')
##            # получаем первый элемент списка удаляя его из исходного списка - Обозначение
##            obozn = name.pop(0)
##            # собираем разделяя пробелами остатки списка в строку - Наименование
##            name = ' '.join(name)
##
##            iPart7 = iKompasDocument3D.TopPart
##            # Меняем свойство Обозначение
##            iPart7.Marking = obozn
##            # Меняем свойство Название
##            iPart7.Name = name
##
##            iModelObject = kompas_api7_module.IModelObject(iPart7)
##            # Без этого обновления изменения не вступят в силу
##            iModelObject.Update()
##
##            iKompasDocument.Save()
##            iKompasDocument.Close(0) # iKompasDocument.Close(1) без iKompasDocument.Save() почему-то не работает
##        iApplication.MessageBoxEx( "Свойства моделей обновлены!", "Сообщение", 64)
##    else:
##        iApplication.MessageBoxEx( "В указанной папке нет 3D-моделей!", "Сообщение", 0)
