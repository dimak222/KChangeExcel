# -*- coding: utf-8 -*-

import pythoncom
from win32com.client import Dispatch, gencache

def getKompasApi():
    module =  gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0)
    return module, api, const.constants

module, api, const =  getKompasApi()
app = api.Application
doc = app.ActiveDocument
doc3D = module.IKompasDocument3D(doc._oleobj_.QueryInterface(module.IKompasDocument3D.CLSID, pythoncom.IID_IDispatch))
propMng = module.IPropertyMng(app._oleobj_.QueryInterface(module.IPropertyMng.CLSID, pythoncom.IID_IDispatch))
propCount = propMng.PropertyCount(doc)
topPart = doc3D.TopPart
parts = topPart.Parts

for i in range(parts.Count):
    print("***************")
    part = module.IPart7(parts.Item(i))
    propKeeper = module.IPropertyKeeper(part._oleobj_.QueryInterface(module.IPropertyKeeper.CLSID, pythoncom.IID_IDispatch))

    for j in range(propCount):
        prop = propMng.GetProperty(doc, j)
        value, fromSource = "", True
        res, value, fromSource = propKeeper.GetPropertyValue(prop, value, True, fromSource)

        if value == None:
            continue

        if isinstance(value, bytes):
            if value != "":
                print(prop.Name + " = " + value)
        else:
           print(j, prop.Name, value)