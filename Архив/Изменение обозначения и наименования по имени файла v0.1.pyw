# -*- coding: cp1251 -*-
# ��������� ���������� ������� 3d ������� �� ��� �� ������

import pythoncom, os, openpyxl
from win32com.client import Dispatch, gencache
try:
    # ������ ������ ��� python 2.xx
    from tkFileDialog import *
except:
    # ������ ������ ��� python 3.xx
    from tkinter.filedialog import *

kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
kompas_api_object = kompas_api7_module.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
iApplication = kompas_api_object.Application

# ������ �������� �����
directory = askdirectory()
if directory:
    # �������� ������ ���� ������ ��������� � �����
    all_files = os.listdir(directory)
    # ��������� �����, ������� ������ 3D-������
    files = filter(lambda x: x.endswith('.m3d'), all_files)

wb = openpyxl.reader.excel.load_workbook(filename="��������� ����������� � ������������ v0.1.xlsx", data_only=True)
wb.active = 0
sheet = wb.active
print(sheet['�18'].value)

    if files:
        # ������ ���� ��� ������ � ������ 3D-�������
        for f in files:
            # ��������� ������ ���� � �����
            PathName = directory + '/' + f
            iDocuments = iApplication.Documents
            print(PathName)

            # ��������� ���� ��������
##            iKompasDocument = iDocuments.Open (PathName, False, False )# False - � ��������� ������, False - � ������������ ��������������
##            iKompasDocument3D = kompas_api7_module.IKompasDocument3D(iKompasDocument)
##
##            # � ��������� ������ iKompasDocument.Name ���������� ������ ����, ������� �������� ������
##            name = iKompasDocument.Name.split('/')[-1]
##            # ���� ����� ��� ��������� 4-x ��������
##            name = name[:-4]
##            # ��������� �� ������ �� ��������
##            name = name.split('_')
##            # �������� ������ ������� ������ ������ ��� �� ��������� ������ - �����������
##            obozn = name.pop(0)
##            # �������� �������� ��������� ������� ������ � ������ - ������������
##            name = ' '.join(name)
##
##            iPart7 = iKompasDocument3D.TopPart
##            # ������ �������� �����������
##            iPart7.Marking = obozn
##            # ������ �������� ��������
##            iPart7.Name = name
##
##            iModelObject = kompas_api7_module.IModelObject(iPart7)
##            # ��� ����� ���������� ��������� �� ������� � ����
##            iModelObject.Update()
##
##            iKompasDocument.Save()
##            iKompasDocument.Close(0) # iKompasDocument.Close(1) ��� iKompasDocument.Save() ������-�� �� ��������
##        iApplication.MessageBoxEx( "�������� ������� ���������!", "���������", 64)
##    else:
##        iApplication.MessageBoxEx( "� ��������� ����� ��� 3D-�������!", "���������", 0)
