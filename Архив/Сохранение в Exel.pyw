#-------------------------------------------------------------------------------
# Name:        Ссылки в Exel
# Purpose:     Запись ссылки в Exel
#
# Author:      dimak
#
# Created:     09.10.2021
# Copyright:   (c) dimak 2021
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import win32com.client

excel = r'D:\YandexDisk\Работа\ГОСТ\Макросы\Изменение обозначения и наименования по имени файла\Тест.xlsx'

xlApp = win32com.client.Dispatch("Excel.Application")
workbook = xlApp.Workbooks.Open(excel)
worksheet = workbook.Worksheets("Лист1")

for xlRow in range(3, 11, 1):
    worksheet.Hyperlinks.Add(Anchor = worksheet.Range('B{}'.format(xlRow)),
    Address="http://www.microsoft.com",
    ScreenTip="Microsoft Web Site",
    TextToDisplay="Microsoft")
workbook.Save()
#workbook.Close()