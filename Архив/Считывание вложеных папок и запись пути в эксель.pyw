#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      dimak222
#
# Created:     16.10.2021
# Copyright:   (c) dimak222 2021
# Licence:     No
#-------------------------------------------------------------------------------

import glob
import pythoncom, os
import win32com.client
from win32com.client import Dispatch, gencache
from tkinter.filedialog import *

kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
kompas_api_object = kompas_api7_module.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
iApplication = kompas_api_object.Application
iDocuments = iApplication.Documents

directory = askdirectory()

#excel = "Тест.xlsx"

xlApp = win32com.client.Dispatch("Excel.Application")
workbook = xlApp.Workbooks.Open(os.path.abspath('Тест.xlsx'))
worksheet = workbook.Worksheets("Лист1")

file_cnt = 0
Row = 2
for name in glob.glob(directory+"/**/*.?3d", recursive=True):
    print(name)
    iKompasDocument = iDocuments.Open (name, False, False )# False - в невидимом режиме, False - с возможностью редактирования
    iKompasDocument3D = kompas_api7_module.IKompasDocument3D(iKompasDocument)
    iPart7 = iKompasDocument3D.TopPart
    Row = Row + 1
    worksheet.Hyperlinks.Add(Anchor = worksheet.Range('B{}'.format(Row)),
    Address = name,
    ScreenTip = iPart7.Marking,
    TextToDisplay = iPart7.Marking)
    #iKompasDocument.Save()
    iKompasDocument.Close(0) # iKompasDocument.Close(1) без iKompasDocument.Save() почему-то не работает
    file_cnt = file_cnt + 1
end_msg = "Свойства моделей обновлены! " "Всего файлов изменено: " + str(file_cnt)
workbook.Close(True)
#open('Тест.xlsx')
print(end_msg)
##for name2 in glob.glob(directory+"/**/*.exe", recursive=True):
##    #name2 = ''.join(name2)
##    print(name2)

##for f in name1:
##    f = open(name1, 'r')
##    print(f.read())
##    f.close()  # не забывайте закрыть файл


### Получаем список всех файлов имеющихся в папке
##all_files = os.listdir(directory)
### Фильтруем файлы, отбирая только 3D-модели и сборки
##files1 = filter(lambda x: x.endswith('.txt'), all_files)
##files2 = filter(lambda x: x.endswith('.exe'), all_files)
##files = list(files1) + list(files2)
##print(files)
