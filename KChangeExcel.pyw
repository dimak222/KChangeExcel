#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     15.03.2023
# Copyright:   (c) dimak222 2023
# Licence:     No
#-------------------------------------------------------------------------------

title = "KChangeExcel"
ver = "v0.1.2.0"
url = "https://github.com/dimak222/KChangeExcel" # ссылка на файл

#------------------------------Настройки!---------------------------------------
property_library = r"C:\Users\Каширских Дмитрий\Desktop\Дмитрий\ГОСТ\Прочее\Макросы\KChangeExcel\Тест\Свойства СПМ.lpt"
#-------------------------------------------------------------------------------

def DoubleExe():# проверка на уже запущеное приложение

    import os # работа с файовой системой
    import psutil # модуль вывода запущеных процессов

    global program_directory # значение делаем глобальным

    filename = title + ".exe" # имя запускаемого файла

    list = [] # список найденых путей

    for process in psutil.process_iter(): # перебор всех процессов

        try: # попытаться узнать имя процесса
            proc_name = process.name() # имя процесса

        except psutil.NoSuchProcess: # в случае ошибки
            pass # пропускаем

        else: # если есть имя
            if proc_name == filename: # сравниваем имя
                list.append(process.cwd())
                if len(list) > 2: # если найдено больше двух названий программы (два процесса)
                    Message("Приложение уже запущено!") # сообщение, поверх всех окон и с автоматическим закрытием
                    exit() # выходим из програмы

    if list == []: # если путь не указан
        program_directory = "" # ничего не указывааем (будет рядом с программой)
    else: # если путь найден
        program_directory = os.path.dirname(list[0]) # путь к файлу

def Check_update(): # проверить обновление приложение

    try: # попытаться импортировать модуль обновления

        import sys # модуль информации об операционной системе

        sys.path.insert(0, "../Updater") # путь откуда брать модуль

        from Updater import Update # импортируем модуль обновления

        Update(title, ver, beta, url) # проверяем обновление (имя программы, версия программы, скачивать бета версию, ссылка на программу)

    except: # не удалось
        pass # пропустить

def KompasAPI(): # подключение API КОМПАСа

    import pythoncom # модуль для запуска без IDE
    from win32com.client import Dispatch, gencache # библиотека API Windows
    from sys import exit # для выхода из приложения без ошибки

    try: # попытаться подключиться к КОМПАСу

##        global KompasConst # значение делаем глобальным
        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
##        global iKompasObject # значение делаем глобальным
        global iKompasDocument # значение делаем глобальным
        global iDocuments # значение делаем глобальным

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
        KompasConst2D = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа 2D документов
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения

        KompasAPI5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch("Kompas.Application.5", None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch("Kompas.Application.7") # интерфейс приложения КОМПАС-3D.

        iKompasDocument = iApplication.ActiveDocument # получить текущий активный документ

        iDocuments = iApplication.Documents # интерфейс для открытия документов

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось подключиться к КОМПАСу

        Message("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def Kompas_message(text): # сообщение в окне КОМПАСа если он открыт

    if iApplication.Visible == True: # если компас видимый
        iApplication.MessageBoxEx(text, 'Message:', 64) # сообщение в КОМПАСе

def Message(text = "Ошибка!", counter = 4): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    from threading import Thread # библиотека потоков

    def Resource_path(relative_path): # для сохранения картинки внутри exe файла

        import os # работа с файовой системой

        try: # попытаться определить путь к папке
            base_path = sys._MEIPASS # путь к временной папки PyInstaller

        except Exception: # если ошибка
            base_path = os.path.abspath(".") # абсолютный путь

        return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

    def Message_Thread(text, counter): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        import tkinter.messagebox as mb # окно с сообщением
        import tkinter as tk # модуль окон

        if counter == 0: # время до закрытия окна (если 0)
            counter = 1 # закрытие через 1 сек
        window_msg = tk.Tk() # создание окна
        try: # попытаться использовать значёк
            window_msg.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        except: # если ошибка
            pass # пропустить
        window_msg.attributes("-topmost",True) # окно поверх всех окон
        window_msg.withdraw() # скрываем окно "невидимое"
        time = counter * 1000 # время в милисекундах
        window_msg.after(time, window_msg.destroy) # закрытие окна через n милисекунд
        if mb.showinfo(title, text, parent = window_msg) == "": # информационное окно закрытое по времени
            pass # пропустить
        else: # если не закрыто по времени
            window_msg.destroy() # окно закрыто по кнопке
        window_msg.mainloop() # отображение окна

    msg_th = Thread(target = Message_Thread, args = (text, counter)) # запуск окна в отдельном потоке
    msg_th.start() # запуск потока

    msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

def Connecting_to_Excel(): # подключение EXcel и изменение св-в

    import os # работа с файовой системой
    from openpyxl import load_workbook # модуль управлением Excel

    name_txt_file = program_directory + title + ".xlsx"

    if os.path.exists(name_txt_file): # если есть txt файл использовать его

        workbook = load_workbook(name_txt_file) # подключимся к Excel

        try: # попытаться подключиться к файлу Excel

            ws_name = "v1.0" # название листа
            ws = workbook[ws_name] # выберем лист

        except: # если файл Excel не найден

            Message("Лист в файле Excel отсутствует или имеет несоответствующию версию!", 8) # сообщение, поверх всех окон с автоматическим закрытием
            exit() # выходим из програмы

        for row in ws.iter_rows(min_col = 2, min_row = 2, max_col = ws.max_column, max_row = ws.max_row, values_only = True): # обойдём все строки и колонки Excel

            if row[0] != None or row[1] != None: # если есть запись старого обозначения
                Change_property(row, ws) # изменение св-в документов (кортеж значений, лист Excel)

##            elif row[1] != None: # если есть запись старого наименования
##                Change_property(row) # изменение св-в документов

    else:
        Message("Файл Excel не найден: " + name_txt_file + "\n Положите его рядом с программой!", 8) # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def Change_property(row, ws): # изменение св-в документов (кортеж значений, лист Excel)

        iKompasDocument3D = KompasAPI7.IKompasDocument3D(iKompasDocument) # базовый класс документов-моделей КОМПАС
        iPart7 = iKompasDocument3D.TopPart # интерфейс компонента 3D документа

        Marking_old = row[0] # старое обозначение
        iName_old = row[1] # старое наименование

        Marking = iPart7.Marking # обозначение документа
        iName = iPart7.Name # наименование документа

        if Marking == Marking_old or Marking_old == None and iName == iName_old: # если обозначение или наименование пустое и наименование совпало

            iPropertyMng = KompasAPI7.IPropertyMng(iApplication) # интерфейс Менеджера свойств

            iPropertyKeeper = KompasAPI7.IPropertyKeeper(iPart7) # интерфейс получения/редактирования значения свойств

            if iPart7.Detail: # если дет.

                for n in range(4, len(row) + 1): # обработка всех столбцов
                    print(n)
                    cell = ws.cell(row = 1, column = n)
                    cell = cell.value
                    print(cell)

                    iProperty = iPropertyMng.GetProperty(iKompasDocument, cell) # интерфейс свойства
                    iPropertyValue = iPropertyKeeper.GetPropertyValue(iProperty, 0, True)[1] # получить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
##                    print(iPropertyValue)

    ##            iGetProperties = iPropertyMng.GetProperties(iKompasDocument) # получить массив св-в
    ##
    ##            for iProperty in iGetProperties:
    ##
    ##                iPropertyValue = iPropertyKeeper.GetPropertyValue(iProperty, 0, True)[1] # получить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
    ##                print(iProperty.Id, iProperty.Name, iPropertyValue)

            ##    iProperty = iPropertyMng.GetProperty(iKompasDocument, "Раздел спецификации") # интерфейс свойства
            ##    iProperty = iPropertyMng.GetProperty(iKompasDocument, "S") # интерфейс свойства

                if iProperty == None:

        ##            iGetProperties = iPropertyMng.GetProperty(property_library, "S") # интерфейс свойства
        ##            iProperty = iPropertyMng.AddProperty(iKompasDocument, iGetProperties) # создаём св-во
        ##            iProperty.Update() # применим сво-ва
                    print("Необходимо создать св-во!")

            else: # если это СБ
                print("Это СБ!")

        else: # пропускаем
            pass

##    iIsComplexPropertyValue = iPropertyKeeper.IsComplexPropertyValue(iProperty) # признак комплексного значения свойства - True - составное свойство, False - нет
##
##    dict_xml = {"Детали": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">' # словарь, где значение до ":" это ключ, после значение.
##                              '<property id="sectionName" value="Детали" type="string" />'
##                              '<property id="sectionNumb" value="20" type="int" />',
##                "Стандартные изделия": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">'
##                                          '<property id="sectionName" value="Стандартные изделия" type="string" />'
##                                          '<property id="sectionNumb" value="25" type="int" />',
##                "Прочие изделия": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">'
##                                      '<property id="sectionName" value="Прочие изделия" type="string" />'
##                                      '<property id="sectionNumb" value="30" type="int" />',
##                "Материалы": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">'
##                                  '<property id="sectionName" value="Материалы" type="string" />'
##                                  '<property id="sectionNumb" value="35" type="int" />',
##                "Комплекты": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">'
##                                  '<property id="sectionName" value="Комплекты" type="string" />'
##                                  '<property id="sectionNumb" value="40" type="int" />',
##                }
##
##    iSetComplexPropertyValue = iPropertyKeeper.SetComplexPropertyValue(iProperty, dict_xml["Детали"]) # установить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
##
##    iProperty = iPropertyMng.GetProperty(iKompasDocument, "Плотность") # интерфейс свойства
##    ##iProperty = iPropertyMng.GetProperty(iKompasDocument, "Материал") # интерфейс свойства
##
##    iPropertyValue = iPropertyKeeper.GetPropertyValue(iProperty, 0, True)[1] # получить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
##    iSetPropertyValue = iPropertyKeeper.SetPropertyValue(iProperty, "2770", True) # получить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
##    ##iSetPropertyValue = iPropertyKeeper.SetPropertyValue(iProperty, "Сталь 40Х ГОСТ 4543-2016", True) # получить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
##
##    iProperty.Update() # применим сво-ва

#-------------------------------------------------------------------------------

DoubleExe() # проверка на уже запущеное приложени

Check_update() # проверить обновление приложение

KompasAPI() # подключение API компаса

if iKompasDocument: # проверяем открыт ли файл в КОМПАСе

    Connecting_to_Excel() # подключение EXcel и изменение св-в

else: # файл не открыт в КОМПАСе
    Message("Откройте дет. или СБ!") # сообщение, поверх всех окон с автоматическим закрытием