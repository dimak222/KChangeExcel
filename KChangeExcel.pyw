#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     15.03.2023
# Copyright:   (c) dimak222 2023
# Licence:     No
#-------------------------------------------------------------------------------

title = "KChangeExcel"
ver = "v0.6.0.0"
url = "https://github.com/dimak222/KChangeExcel" # ссылка на файл

#------------------------------Настройки!---------------------------------------
check_update = False # проверять обновление программы ("True" - да; "False" или "" - нет)
beta = False # скачивать бета версии программы ("True" - да; "False" или "" - нет)

property_library = "Свойства СПМ.lpt"
#-------------------------------------------------------------------------------

from sys import exit # для выхода из приложения без ошибки

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
                    exit() # выходим из программы

    if list == []: # если путь не указан
        program_directory = "" # ничего не указывааем (будет рядом с программой)
    else: # если путь найден
        program_directory = os.path.abspath(list[0]) # путь к файлу

def Check_update(): # проверить обновление приложение

    if check_update:

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

        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
        global iKompasObject # значение делаем глобальным
        global iDocuments # значение делаем глобальным

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
        KompasConst2D = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа 2D документов
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения

        KompasAPI5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch("Kompas.Application.5", None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch("Kompas.Application.7") # интерфейс приложения КОМПАС-3D.

        iDocuments = iApplication.Documents # интерфейс для открытия документов

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось подключиться к КОМПАСу

        Message("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из программы

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

def SystemPath(): # определяем папку системных файлов

    global iSystemPath # значение делаем глобальным

    iSystemPath = iKompasObject.ksSystemPath(0) # определяем папку системных файлов

def Сhecking_open_files(): # проверка открытых файлов

    from sys import exit # для выхода из приложения без ошибки

    iKompasDocument = iApplication.ActiveDocument # получить текущий активный документ

    if iKompasDocument == None or iKompasDocument.DocumentType not in (4, 5, 7): # если не открыт док. или не 2D док., выдать сообщение (1-чертёж; 2- фрагмент; 3-СП; 4-модель; 5-СБ; 6-текст. док.; 7-тех. СБ;)

        Message("Откройте дет. или СБ!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из программы

def Connecting_to_Excel(): # подключение EXcel и изменение св-в

    import os # работа с файовой системой
    from openpyxl import load_workbook # модуль управлением Excel

    global name_txt_file # значение делаем глобальным
    global workbook  # значение делаем глобальным
    global ws # значение делаем глобальным

    name_txt_file = os.path.join(program_directory, title + ".xlsx")

    if os.path.exists(name_txt_file): # если есть txt файл использовать его

        workbook = load_workbook(name_txt_file) # подключимся к Excel

        try: # попытаться подключиться к файлу Excel

            ws_name = "v1.0" # название листа
            ws = workbook[ws_name] # выберем лист

        except: # если файл Excel не найден

            Message("Лист в файле Excel отсутствует или имеет несоответствующию версию!", 8) # сообщение, поверх всех окон с автоматическим закрытием
            exit() # выходим из программы

        for row in ws.iter_rows(min_col = 2, min_row = 2, max_col = ws.max_column, max_row = ws.max_row, values_only = True): # обойдём все строки и колонки Excel

            if row[0] != None or row[1] != None: # если есть запись старого обозначения или наименования
                list_Excel.append(list(row)) # записать колонку в список

    else:
        Message("Файл Excel не найден: " + name_txt_file + "\n Положите его рядом с программой!", 8) # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из программы

def Main_assembly(): # обработка главной СБ или дет. (список значений)

    iKompasDocument = iApplication.ActiveDocument # получить текущий активный документ

    iKompasDocument3D = KompasAPI7.IKompasDocument3D(iKompasDocument) # базовый класс документов-моделей КОМПАС
    iPart7 = iKompasDocument3D.TopPart # интерфейс компонента 3D документа (сам документ)

    parameters = iPart7.Marking, iPart7.Name, iPart7.FileName # список параметров файла

    list_files.append(parameters) # добавляем в список

    if iPart7.Detail: # если это дет.

        Сhecking_match(False) # проверка на совпадение обозначения или наименования (не открывать/закрывать файл)

    else: # если это СБ

        Сhecking_match(False) # проверка на совпадение обозначения или наименования (не открывать/закрывать файл)

        Collect_sources(iPart7) # рекурсивный сбор дет. и СБ (интерфейс компонента 3D документа)

        Сhecking_match(True) # проверка на совпадение обозначения или наименования (не открывать/закрывать файл)

        iKompasDocument3D.RebuildDocument() # перестроить СБ
        iKompasDocument3D.Save() # сохранить изменения

def Collect_sources(iPart7): # рекурсивный сбор дет. и СБ

    try: # попытаться найти все елементы в СБ

        iPartsEx = iPart7.PartsEx(1) # список компонентов, включённыхв расчёт (0 - все компоненты (включая копии из операций копирования); 1 - первые экземпляры вставок компонентов (ksPart7CollectionTypeEnum))

        for iPart7 in iPartsEx: # проверяем каждый элемент из вставленных в СБ

            if iPart7.Standard: # если это стандартная дет.
                print("Стандартная дет.!", iPart7.FileName)
                continue

            elif iPart7.IsLocal: # если это локальная СБ пропускаем её
                print("Локальная сборка!", iPart7.FileName)
                continue

            elif iPart7.IsBillet: # если СБ заготовка пропускаем её
                print("Вставлена заготовка!", iPart7.FileName)
                continue

            elif iPart7.IsLayoutGeometry: # если это компоновочная геометрия
                print("Компоновочная геометрия!", iPart7.FileName)
                continue

            parameters = iPart7.Marking, iPart7.Name, iPart7.FileName # список параметров файла

            list_files.append(parameters) # добавляем в список

            if not iPart7.Detail: # если это СБ
                Collect_Sources(iPart7) # рекурсивный сбор уникальных документов

    except: # если ошибка (нет дет.)
        print("Пустая сборка!")
        pass

def Сhecking_match(iClose): # проверка на совпадение обозначения или наименования (не открывать/закрывать файл)

    for row in list_Excel: # проверяем каждую строчку считаную с Excel

        iMarking_old = row[0] # старое обозначение с удалением пробелов по бокам
        iName_old = row[1] # старое наименование с удалением пробелов по бокам

        for file in list_files: # проверяем каждого файла на совпадение

            iMarking = file[0] # обозначение документа с удалением пробелов по бокам
            iName = file[1] # наименование документа с удалением пробелов по бокам

            if iMarking == iMarking_old or iMarking_old == None and iName == iName_old: # если обозначение или наименование пустое и наименование совпало

                if iClose: # не открывать/закрывать файл
                    iKompasDocument = iDocuments.Open(file[2], False, False) # Открытие файлов (False - в невидимом режиме, False - с возможностью редактирования)

                else: # получить тукущий активный докумен
                    iKompasDocument = iApplication.ActiveDocument # получить текущий активный докумен

                if Сhange_properties(row, file, iKompasDocument): # изменение св-в документов (список значений, параметры считаных файлов)

                    iKompasDocument.Save() # iKompasDocument.Close(1) без iKompasDocument.Save() почему-то не работает

                    if iClose: # не открывать/закрывать файл
                        iKompasDocument.Close(1) # 0 - закрыть документ без сохранения; 1 - закрыть документ, сохранив  изменения; 2 - выдать запрос на сохранение документа, если он изменен.

                if not iClose: # не открывать/закрывать файл
                    list_files.pop(0) # удаляем запись из списка

def Сhange_properties(row, file, iKompasDocument): # изменение св-в документов (список значений, параметры считаных файлов)

    import os # работа с файовой системой

    Сhange = False # тригер изменения

    iKompasDocument3D = KompasAPI7.IKompasDocument3D(iKompasDocument) # базовый класс документов-моделей КОМПАС
    iPart7 = iKompasDocument3D.TopPart # интерфейс компонента 3D документа (сам документ)

    iPropertyMng = KompasAPI7.IPropertyMng(iApplication) # интерфейс менеджера свойств
    iPropertyKeeper = KompasAPI7.IPropertyKeeper(iPart7) # интерфейс получения/редактирования значения свойств

    for n in range(0, len(row) - 3): # обработка всех столбцов кроме последнего

        cell_name = ws.cell(row = 1, column = n + 4) # колонки с заголовками с пропуском первых 3-х
        cell_name = str(cell_name.value).strip() # значение колонок

        cell = str(row[n+2]).strip() # значение из списка

        if cell == "None": # если значение ячейки не записанно
            continue # пропустить

        iGetProperties = iPropertyMng.GetProperties(iKompasDocument) # получить массив св-в

        for iProperty in iGetProperties: # перебор всех св-в из массива

            if cell_name == iProperty.Name.strip(): # сравниваем наименование значения с Excel и наименование значения считаные с документа

                iProperty = iPropertyMng.GetProperty(iKompasDocument, cell_name) # интерфейс свойства
                iPropertyValue = str(iPropertyKeeper.GetPropertyValue(iProperty, 0, True)[1]).strip() # получить значение св-ва (интерфейс св-ва, значение св-ва, единици измерения (СИ))

                if cell != iPropertyValue: # сравниваем значения с Excel и считаные с документа

                    if cell_name == "Раздел спецификации": # если составное св-во

                        dict_xml = {"Сборочные единицы": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">'
                                                         '<property id="sectionName" value="Сборочные единицы" type="string" />'
                                                         '<property id="sectionNumb" value="15" type="int" />',
                                    "Детали": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">' # словарь, где значение до ":" это ключ, после значение.
                                              '<property id="sectionName" value="Детали" type="string" />'
                                              '<property id="sectionNumb" value="20" type="int" />',
                                    "Стандартные изделия": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">'
                                                           '<property id="sectionName" value="Стандартные изделия" type="string" />'
                                                           '<property id="sectionNumb" value="25" type="int" />',
                                    "Прочие изделия": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">'
                                                      '<property id="sectionName" value="Прочие изделия" type="string" />'
                                                      '<property id="sectionNumb" value="30" type="int" />',
                                    "Материалы": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">'
                                                 '<property id="sectionName" value="Материалы" type="string" />'
                                                 '<property id="sectionNumb" value="35" type="int" />',
                                    "Комплекты": '<property id="SPCSection" expression="" fromSource="false" format="{$sectionName}">'
                                                 '<property id="sectionName" value="Комплекты" type="string" />'
                                                 '<property id="sectionNumb" value="40" type="int" />',
                                    }

                        iSetComplexPropertyValue = iPropertyKeeper.SetComplexPropertyValue(iProperty, dict_xml[cell]) # установить значение св-ва (интерфейс св-ва, значение св-ва, единици измерения (СИ))
                        iProperty.Update() # применим сво-ва

                        Сhange = True # тригер изменения

                        print(f"{cell_name}: \"{iPropertyValue}\" => \"{cell}\" изменено!")

                        break # прерываем цикл

                    else:

                        iSetPropertyValue = iPropertyKeeper.SetPropertyValue(iProperty, cell, True) # установить значение св-ва (интерфейс св-ва, значение св-ва, единици измерения (СИ))
                        iProperty.Update() # применим сво-ва

                        Сhange = True # тригер изменения

                        print(f"{cell_name}: \"{iPropertyValue}\" => \"{cell}\" изменено!")

                        break # прерываем цикл

                else: # если свойства совпадают
                    print(f"{cell_name}: \"{cell}\" уже записанно!")
                    break # прерываем цикл

            else: # если свойство не совпало
                continue # использовать следующее

        else: # если св-во не найдено в св-х дет.

            property_library_path = os.path.join(iSystemPath, property_library) # путь к библиотеке со свойствами

            if os.path.exists(property_library_path): # если есть txt файл использовать его

                iGetProperties = iPropertyMng.GetProperty(property_library_path, cell_name) # интерфейс свойства
                iProperty = iPropertyMng.AddProperty(iKompasDocument, iGetProperties) # создаём св-во

                if iProperty != None: # если есть cв-во в библиотеке

                    iPropertyValue = str(iPropertyKeeper.GetPropertyValue(iProperty, 0, True)[1]).strip() # получить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
                    iSetPropertyValue = iPropertyKeeper.SetPropertyValue(iProperty, cell, True) # установить значение свойства (интерфейс св-ва, значение св-ва, единици измерения (СИ))
                    iProperty.Update() # применим сво-ва

                    Сhange = True # тригер изменения

                    print(f"{cell_name}: \"{iPropertyValue}\" => \"{cell}\" изменено и взято из библиотеки!")

                else: # если св-во отсутствует
                    Message(f"Отсутствует св-во \"{cell_name}\" в библиотеке \"{property_library_path}\"") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
                    exit() # выходим из программы

            else:
                Message(f"Не обнаружена библиотека св-в: \"{property_library_path}\"\nРазместите её по указаному пути!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
                exit() # выходим из программы

    return Сhange # тригер изменения

def Change_or_not(n, value): # запись изменения в последнюю колонку

    cell = ws.cell(row = n, column = ws.max_column) # указание колонки
    cell.value = value # запись значения в колонку

    try: # попытаться сохранить записанные значения

        workbook.save(name_txt_file) # сохраняем Excel перед считыванием

    except: # если файл Excel не найден

        print(f"В ячейке \"Изменено\" - \"{n-1}\":{value}")

#-------------------------------------------------------------------------------

list_Excel = [] # список считанных строк
list_files = [] # список считаных файлов

DoubleExe() # проверка на уже запущеное приложени

Check_update() # проверить обновление приложение

KompasAPI() # подключение API компаса

SystemPath() # определяем папку системных файлов

Сhecking_open_files() # проверка открытых файлов

Connecting_to_Excel() # подключение EXcel и изменение св-в

Main_assembly() # обработка главной СБ или дет. (список значений)

##n = 1 # счётчик обработаных строк
##
##for row in list_Excel: # если есть запись старого обозначения
##
##    n += 1 # счётчик обработаных строк
##
##    Change_or_not(n, "Нет") # запись изменения в последнюю колонку
##
##    if Main_assembly(row): # обработка главной СБ или дет. (список значений)
##        Change_or_not(n, "Да") # запись изменения в последнюю колонку
##
##else: # если
##    print("Проверка закончена!")
