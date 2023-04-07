#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     15.03.2023
# Copyright:   (c) dimak222 2023
# Licence:     No
#-------------------------------------------------------------------------------

title = "KChangeExcel"
ver = "v0.8.1.0"
url = "https://github.com/dimak222/KChangeExcel" # ссылка на файл

#------------------------------Настройки!---------------------------------------

check_update = False # проверять обновление программы ("True" - да; "False" или "" - нет)
beta = False # скачивать бета версии программы ("True" - да; "False" или "" - нет)

property_library = "Свойства СПМ.lpt" # название файла со св-ми

#------------------------------Импорт модулей-----------------------------------

import os # работа с файовой системой
import psutil # модуль вывода запущеных процессов

import sys # модуль информации об операционной системе

import pythoncom # модуль для запуска без IDE
from win32com.client import Dispatch, gencache # библиотека API Windows
from sys import exit # для выхода из приложения без ошибки

from threading import Thread # библиотека потоков

import tkinter.messagebox as mb # окно с сообщением
import tkinter as tk # модуль окон

from openpyxl import load_workbook # модуль управлением Excel

import tkinter.ttk as ttk # модуль окон
import time # модуль времени

from pythoncom import VT_EMPTY # для записи пустого св-ва
from win32com.client import VARIANT # для записи пустого св-ва

import re # модуль регулярных выражений

#-------------------------------------------------------------------------------

def DoubleExe(): # проверка на уже запущеное приложение

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

            sys.path.insert(0, "../Updater") # путь откуда брать модуль

            from Updater import Update # импортируем модуль обновления

            Update(title, ver, beta, url) # проверяем обновление (имя программы, версия программы, скачивать бета версию, ссылка на программу)

        except: # не удалось
            pass # пропустить

def KompasAPI(): # подключение API КОМПАСа

    try: # попытаться подключиться к КОМПАСу

        global KompasConst # значение делаем глобальным
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

    def Resource_path(relative_path): # для сохранения картинки внутри exe файла

        try: # попытаться определить путь к папке
            base_path = sys._MEIPASS # путь к временной папки PyInstaller

        except Exception: # если ошибка
            base_path = os.path.abspath(".") # абсолютный путь

        return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

    def Message_Thread(text, counter): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

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

    iKompasDocument = iApplication.ActiveDocument # получить текущий активный документ

    if iKompasDocument == None or iKompasDocument.DocumentType not in (4, 5, 7): # если не открыт док. или не 2D док., выдать сообщение (1-чертёж; 2- фрагмент; 3-СП; 4-модель; 5-СБ; 6-текст. док.; 7-тех. СБ;)

        Message("Откройте дет. или СБ!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из программы

def Connecting_to_Excel(): # подключение EXcel и изменение св-в

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

def Main_assembly(): # обработка главной СБ или дет.

    iKompasDocument = iApplication.ActiveDocument # получить текущий активный документ

    iKompasDocument3D = KompasAPI7.IKompasDocument3D(iKompasDocument) # базовый класс документов-моделей КОМПАС
    iPart7 = iKompasDocument3D.TopPart # интерфейс компонента 3D документа (сам документ)

    iEmbodimentsManager = KompasAPI7.IEmbodimentsManager(iPart7) # интерфейс менеджера исполнений

    iCurrentEmbodimentIndex = iEmbodimentsManager.CurrentEmbodimentIndex # индекс исполнени

    parameters = iPart7.Marking, iPart7.Name, iCurrentEmbodimentIndex, iPart7.FileName # список параметров файла (Обозначение, Наименование, индекс исполнения, путь к файлу)

    list_files.append(parameters) # добавляем в список

    if iPart7.Detail: # если это дет.

        Сhecking_match(False, iKompasDocument) # проверка на совпадение обозначения или наименования (не открывать/закрывать файл)

    else: # если это СБ

        Сhecking_match(False, iKompasDocument) # проверка на совпадение обозначения или наименования (не открывать/закрывать файл)

        Collect_sources(iPart7) # рекурсивный сбор дет. и СБ (интерфейс компонента 3D документа)

        Сhecking_match(True, iKompasDocument) # проверка на совпадение обозначения или наименования (не открывать/закрывать файл)

        iKompasDocument3D.RebuildDocument() # перестроить СБ
        iKompasDocument3D.Save() # сохранить изменения

def Collect_sources(iPart7): # рекурсивный сбор дет. и СБ

    iPartsEx = iPart7.PartsEx(1) # список компонентов, включённыхв расчёт (0 - все компоненты (включая копии из операций копирования); 1 - первые экземпляры вставок компонентов (ksPart7CollectionTypeEnum))

    if iPartsEx: # если есть компоненты в сборке

        for iPart7 in iPartsEx: # проверяем каждый элемент из вставленных в СБ

            if iPart7.Standard: # если это стандартная дет.
                print("Стандартная дет.!", iPart7.FileName)
                continue

            elif iPart7.IsLocal: # если это локальная СБ пропускаем её
                print("Локальный компонент!", iPart7.FileName)
                continue

            elif iPart7.IsBillet: # если это заготовка пропускаем её
                print("Вставлена заготовка!", iPart7.FileName)
                continue

            elif iPart7.IsLayoutGeometry: # если это компоновочная геометрия
                print("Компоновочная геометрия!", iPart7.FileName)
                continue

            iEmbodimentsManager = KompasAPI7.IEmbodimentsManager(iPart7) # интерфейс менеджера исполнений

            iCurrentEmbodimentIndex = iEmbodimentsManager.CurrentEmbodimentIndex # индекс исполнени

            parameters = iPart7.Marking, iPart7.Name, iCurrentEmbodimentIndex, iPart7.FileName # список параметров файла (Обозначение, Наименование, индекс исполнения, путь к файлу)

            if parameters not in list_files: # если нет в списке пути к файлу, добавить (исключает добавление одной и той же детали (одинаковый путь к детали))
                list_files.append(parameters) # добавляем в список

            if not iPart7.Detail: # если это СБ

                Collect_sources(iPart7) # рекурсивный сбор уникальных документов

    else: # если ошибка (нет дет.)
        print("Пустая сборка!")
        pass

def Сhecking_match(iClose, iKompasDocument): # проверка на совпадение обозначения или наименования (не открывать/закрывать файл)

    global file_number # для чтения в потоке
    global current_file_name # для чтения в потоке
    global Stop # для чтения в потоке

    file_number = 0 # отчёт от 0-го файла
    current_file_name = "" # что бы избежать ошибки окна в потоке

    all_failes_number = len(list_Excel) # количество всех строк в списке

    if iClose: # не открывать/закрывать файл
        Message_count(all_failes_number, "Идёт обработка файлов!") # выдача сообщений о количестве файлов (количество всех файлов, сообщение) + file_number (номер обрабатываемого файла) + current_file_name (текущее название файла)

    else: # если не запускаем процесс обработки
        Stop = False # триггер остановки сообщения (для работы сообщений при повторном вызове)

    iApplication.HideMessage = KompasConst.ksHideMessageNo # скрыть сообщение перестроения и не перестраивать

    for row in list_Excel: # проверяем каждую строчку считаную с Excel

        file_number += 1 # отчёт количества обработаных файлов

        Change_or_not(file_number + 1, "Нет") # запись изменения в последнюю колонку

        if Stop == False: # если не нажата кнопка "Отмена" или крестик

            iMarking_old = str(row[0]).strip() # старое обозначение с удалением пробелов по бокам
            iMarking_old = MarkingEmbodimentDocCode(iMarking_old) # кортеж старого обозначения (Обозначение, исполнение, пробел перед кодом док., код док.)
            iMarking_old = iMarking_old[0] + iMarking_old[1] # обозначение с исп.

            iName_old = str(row[1]).strip() # старое наименование с удалением пробелов по бокам

            for file in list_files: # проверяем каждого файла на совпадение

                iMarking = file[0].strip() # обозначение документа с удалением пробелов по бокам
                iMarking = MarkingEmbodimentDocCode(iMarking) # кортеж считанного обозначения (Обозначение, исполнение, пробел перед кодом док., код док.)
                iMarking = iMarking[0] + iMarking[1] # обозначение с исп.

                iName = file[1].strip() # наименование документа с удалением пробелов по бокам

                if iMarking_old == iMarking or iMarking_old == None and iName_old == iName: # если обозначение или наименование пустое и наименование совпало

                    current_file_name = os.path.basename(file[3]) # имя документа с расширением для вывода названия файла в окно сообщений

                    if iClose: # не открывать/закрывать файл
                        iKompasDocument = iDocuments.Open(file[3], False, False) # Открытие файлов (False - в невидимом режиме, False - с возможностью редактирования)

                    else: # не открывать/закрывать файл
                        list_files.pop(0) # удаляем запись из списка

                    if Сhange_properties(row, file[2], iKompasDocument): # изменение св-в документов (список значений из Excel, индекс исполнения файла, интерфейс документа)

                        iKompasDocument.Save() # iKompasDocument.Close(1) без iKompasDocument.Save() почему-то не работает

                        if iClose: # не открывать/закрывать файл
                            iKompasDocument.Close(1) # 0 - закрыть документ без сохранения; 1 - закрыть документ, сохранив  изменения; 2 - выдать запрос на сохранение документа, если он изменен.

                        Change_or_not(file_number + 1, "Да") # запись изменения в последнюю колонку

        else: # если нажали кнопку "Отмена" или крестик
            print("Остановили окном!")
            break # прерываем цикл

    Stop = True # триггер остановки обработки и сообщения

    iApplication.HideMessage = KompasConst.ksShowMessage # показывать сообщение перестроения

def Message_count(all_failes_number, msg = "Идёт обработка файлов!"): # выдача сообщений о количестве файлов (количество всех файлов, сообщение) + file_number (номер обрабатываемого файла) + current_file_name (текущее название файла)

    global Stop # глобальный параметр остановки сообщения

    def Message_count_Thread(all_failes_number, msg): # сообщений о количестве файлов в потоке

        global Stop # глобальный параметр остановки обработки

        class ToolTip(object): # отображает подсказку к виджету

            def __init__(self, widget, text):
                self.widget = widget
                self.text = text
                self.acid = None
                self.tipwindow = None
                self.widget.bind('<Enter>', self.enter)
                self.widget.bind('<Leave>', self.leave)
                self.widget.bind('<ButtonRelease>', self.leave)
                self.widget.bind('<Key>', self.leave)

            def enter(self, event):
                self.schedule()

            def leave(self, event):
                self.unschedule()
                self.hidetip()

            def schedule(self):
                self.unschedule()
                self.acid = self.widget.after(300, self.showtip) # через сколько милисунд отображать подсказку

            def unschedule(self):
                idac = self.acid
                if idac:
                    self.widget.after_cancel(idac)
                self.acid = None

            def showtip(self):
                tw = self.tipwindow = tk.Toplevel(self.widget)
                tw.wm_overrideredirect(1)
                tw.wm_attributes('-topmost', 1) # поверх всех окон
                tw.wm_geometry('+%d+%d' % (self.widget.winfo_rootx(), self.widget.winfo_rooty() + self.widget.winfo_height() + 2))
                tk.Label(tw, text = current_file_name, justify = 'left', bg = '#f2f2f2', relief = 'solid', bd = 1, font = "Verdana 10").pack() # положение, цвет и шрифт текста

            def hidetip(self):
                tw = self.tipwindow
                if tw:
                    tw.destroy()
                self.tipwindow = None

        def Update_text(): # обновление отчёта цифр

            def Updating_text(): # обновление текста

                if Stop: # если триггер остановки обработки и сообщения включён
                    print("Остановил поток!")
                    window.destroy() # закрываем окно

                else: # триггер выключен
                    text.config(text = str(file_number) + "/" + str(all_failes_number)) # обновляем текст
                    text.after(300, Updating_text) # через милисекунды запускаем функцию заново

            Updating_text() # обновление текста

        def Update_progress(): # обновление прогресса

            def Updating_progress(): # обновление прогресса

                percent_file_number = percent_all_failes_number * file_number # процент выполнения

                if Stop: # если триггер остановки обработки и сообщения включён
                    print("Остановил поток прогресса!")
                    window.destroy() # закрываем окно

                else: # триггер выключен
                    progress['value'] = percent_file_number # процент выполнения
                    window.update() # (update_idletasks не сбрасывет дпока не дошёл до конца)
                    progress.after(300, Updating_progress) # через милисекунды запускаем функцию заново

            Updating_progress() # обновление прогресса

        def Button_exit(): # кнопка "Отмена"
            window.destroy() # закрываем окно

        window = tk.Tk() # создание окна
        window.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        window.title(title) # заголовок окна
        window.attributes("-topmost",True) # окно поверх всех окон
        x = (window.winfo_screenwidth() - window.winfo_reqwidth()) / 2 # положение по центру монитора
        y = (window.winfo_screenheight() - window.winfo_reqheight()) / 2 # положение по центру монитора
        window.wm_geometry("+%d+%d" % (x-50, y)) # положение по центру монитора -50 из-за логотипа
##        window.geometry('200x100') # размер окна
        window.resizable(width = False, height = False) # блокировка изменение размера окна

##        logo = tk.PhotoImage(file = Resource_path("cat.png")) # логотип
##        logo = logo.subsample(1, 1) # мастаб картинки
##        tk.Label(window, image=logo).pack(side="right") # расположение картинки в окне

        f_top = tk.Frame(window) # блок окна (вверх)
        f_top.pack(expand = True, fill = "both") # размещение блока (с возможностью расширяться и заполненем окна во всех направлениях)

        text = tk.Label(f_top, justify=tk.LEFT, font = "Verdana 10", text = msg) # текст в окне
        text.pack(padx = 5, pady = 2) # размещение блока

        text = tk.Label(f_top, fg="green", justify=tk.LEFT, padx = 3, pady = 3, font = "Verdana 10") # текст
        ToolTip(text, current_file_name) # имя текущего файла в виде всплывающего окна
        Update_text() # обновление отчёта цифр
        text.pack() # размещение блока

        progress = ttk.Progressbar(f_top, orient = "horizontal", length = 250, mode = 'determinate') # панель прогресса (положение, длина, вид отображения)
        percent_all_failes_number = 100/all_failes_number # перевод в процент от общего числа
        Update_progress() # обновление прогресса
        progress.pack(padx = 4) # размещение блока

        button = tk.Button(f_top, font = "Verdana 11", command = Button_exit, text = "Отмена") # действие кнопки
        button.pack(side = "bottom", pady = 3) # размещение блока

        window.mainloop() # отображение окна

        Stop = True # триггер остановки обработки и сообщения

    def Resource_path(relative_path): # для сохранения картинки внутри exe файла

        try: # попытаться определить путь к папке
            base_path = sys._MEIPASS # путь к временной папки PyInstaller

        except Exception: # если ошибка
            base_path = os.path.abspath(".") # абсолютный путь

        return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

    Stop = False # триггер остановки сообщения (для работы сообщений при повторном вызове)

    msg_th = Thread(target = Message_count_Thread, args = (all_failes_number, msg, )) # запуск сообщений о количестве файлов в отдельном потоке
    msg_th.start() # запуск потока

    return msg_th # возращаем запущеный поток для определения его завершения

def Change_or_not(n, value): # запись изменения в последнюю колонку

    cell = ws.cell(row = n, column = ws.max_column) # указание колонки
    cell.value = value # запись значения в колонку

    try: # попытаться сохранить записанные значения

        workbook.save(name_txt_file) # сохраняем Excel перед считыванием

    except: # если файл Excel не найден

        print(f"В ячейке \"Изменено\" - \"{n-1}\":{value}")

def MarkingEmbodimentDocCode(iMarking): # выделение обозначения, исполнения и кода документа, возвращает кортеж (Обозначение, исполнение, пробел перед кодом док., код док.)

    if iMarking != VARIANT(VT_EMPTY, None): # если не надо удалять Обозначение

        docCode =["СБ", "ГЧ", "УЧ", "ЭСБ", "МЧ", "МД", "МС"] # список кодов документа

        for docCode in docCode: # перебор кода документа

            whitespaceDocCode = " " + docCode # добавляем пробел перед кодом

            if iMarking.find(whitespaceDocCode)!= -1: # если код документа с пробелом найден
                iMarking = iMarking.replace(whitespaceDocCode, "", 1) # заменить код документа с пробелом на ""
                whitespaceDocCode = " " # пробел перед кодом документа
                break # прервать цикл

            elif iMarking.find(docCode)!= -1: # если код документа найден
                iMarking = iMarking.replace(docCode, "", 1) # заменить код документа на ""
                whitespaceDocCode = "" # без пробела перед кодом документа
                break # прервать цикл

            else: # код документа не подощёл
                continue # использовать следующий

        else: # код документа не найден
            whitespaceDocCode = " " # пробел перед кодом докумена
            docCode = "" # без кода документа

        embodiment = re.findall("-\d{2,3}$", iMarking, re.M) # определение исп. в конце обозначения

        if embodiment: # если найдено исполнение
            embodiment = embodiment[0] # выбираем первый элемент
            iMarking = iMarking.replace(embodiment, "", 1) # заменить исп. на ""
            embodiment = embodiment.replace("-", "",1) # убрать из исп. "-"

        else: # нет исп.
            embodiment = "" # не пишем исп.

    else: # удаляем Обозначение

        iMarking = "" # обозначение
        embodiment = "" # исп.
        whitespaceDocCode = "" # пробел перед кодом документа
        docCode = "" # код докумета

    return iMarking, embodiment, whitespaceDocCode, docCode # возвращаем значения

def Сhange_properties(row, embodimentIndex, iKompasDocument): # изменение св-в документов (список значений из Excel, индекс исполнения файла, интерфейс документа)

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
            continue # использовать следующее значение

        elif cell == "delete": # если значение ячейки надо удалить

            if cell_name == "Плотность": # если значение плотности удалять
                cell = None # вписать "0"

            elif cell_name == "Наименование": # если значение Наименования удалять

                iPart7.Name = "" # прописываем пустоту
                iPart7.Update() # применяем св-ва

                Сhange = True # тригер изменения

                continue # использовать следующее значение

            else: # все остальнвые
                cell = VARIANT(VT_EMPTY, None) # пропишем пустое значение

        iGetProperties = iPropertyMng.GetProperties(iKompasDocument) # получить массив св-в

        for iProperty in iGetProperties: # перебор всех св-в из массива

            if cell_name == iProperty.Name.strip(): # сравниваем наименование ячейки с Excel и наименование значения считаные с документа

                iProperty = iPropertyMng.GetProperty(iKompasDocument, cell_name) # интерфейс свойства
                iPropertyValue = str(iPropertyKeeper.GetPropertyValue(iProperty, 0, True)[1]).strip() # получить значение св-ва (интерфейс св-ва, значение св-ва, единици измерения (СИ))

                if cell != iPropertyValue: # сравниваем значения с Excel и считаные с документа

                    if cell_name == "Обозначение": # если название ячейки "Обозначение"

                        СheckEmbodiment(iPart7, embodimentIndex) # проверка и зменение исп.

                        tuple_marking = MarkingEmbodimentDocCode(cell) # выделение обозначения, исполнения и кода документа (возвращает кортеж)

                        RecordMarking(tuple_marking, iPart7) # запись обозначения, исполнения и кода документа

                        Сhange = True # тригер изменения

                        break # прерываем цикл

                    elif cell_name == "Раздел спецификации" and not VARIANT(VT_EMPTY, None): # если изменение раздела спецификаци и не удалять значение

                        Сhange_property_SP(iPropertyKeeper, iProperty, cell) # изменение св-в раздела спецификации (интерфейс получения/редактирования значения свойств, интерфейс свойства, значение ячейки)

                        Сhange = True # тригер изменения

                        print(f"{cell_name}: \"{iPropertyValue}\" => \"{cell}\" изменено!")

                        break # прерываем цикл

                    else: # другие св-ва

                        iSetPropertyValue = iPropertyKeeper.SetPropertyValue(iProperty, cell, True) # установить значение св-ва (интерфейс св-ва, значение св-ва, единици измерения (СИ))
                        iProperty.Update() # применим сво-ва

                        Сhange = True # тригер изменения

                        print(f"{cell_name}: \"{iPropertyValue}\" => \"{cell}\" изменено!")

                        break # прерываем цикл

                else: # если свойства совпадают
##                    print(f"{cell_name}: \"{cell}\" уже записанно!")
                    break # прерываем цикл

            else: # если свойство не совпало
                continue # использовать следующее значение

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

def СheckEmbodiment(iPart7, embodimentIndex): # проверка и изменение исп.

    iEmbodimentsManager = KompasAPI7.IEmbodimentsManager(iPart7) # интерфейс менеджера исполнений

    iCurrentEmbodimentIndex = iEmbodimentsManager.CurrentEmbodimentIndex # получаем индекс текущего исполнения

    if iCurrentEmbodimentIndex != embodimentIndex: # если текущее исп. отличаеться от заданного
        iSetCurrentEmbodiment = iEmbodimentsManager.SetCurrentEmbodiment(embodimentIndex) # устанавливаем исп.

def СhangeEmbodiment(obozn, name, var, docCode, whitespacedocCode): # сортировка исполнения (для последовательной записи), присвоение обозначения, наименования, исполнения и кода документа

    iKompasDocument3D = KompasAPI7.IKompasDocument3D(iKompasDocument)
    iEmbodimentsManager = KompasAPI7.IEmbodimentsManager(iKompasDocument3D)

    isp = 0
    ispCount = iEmbodimentsManager.EmbodimentCount                              # узнать количество исполнений
    if ispCount >1:
        Currentisp = iEmbodimentsManager.GetCurrentEmbodimentMarking(2,False)   # узнаю номер текущего исполнения (-1 - всё обозначение; 1 - базовая часть обозначения; 2 - исполнение с прочерком; 3 - "1" и "2" вместе; 8 - код документа с прочерком)
        if Currentisp != "":
            if var == "" or abs(int(Currentisp)) >= int(var):
                while isp < ispCount:
                    record(obozn, name, var, docCode, isp, whitespacedocCode)
                    if var == "":
                        var = "00"
                    var = int(var) + 1
                    if var < 10:
                        var = "0" + str(var)
                    var = str(var)
                    isp = isp + 1
            else:
                isp = isp - 1 + ispCount
                var = int(var) - 1 + ispCount
                if var < 10:
                    var = "0" + str(var)
                var = str(var)
                while isp + 1 > 0:
                    record(obozn, name, var, docCode, isp, whitespacedocCode)
                    var = int(var) - 1
                    if var < 10:
                        var = "0" + str(var)
                    var = str(var)
                    isp = isp - 1
        else:
            isp = isp - 1 + ispCount
            if var == "":
                var = "00"
            var = int(var) - 1 + ispCount
            if var < 10:
                var = "0" + str(var)
            var = str(var)
            while isp + 1 > 0:
                if var == "00":
                    var = ""
                record(obozn, name, var, docCode, isp, whitespacedocCode)
                if var == "":
                    break
                var = int(var) - 1
                if var < 10:
                    var = "0" + str(var)
                var = str(var)
                isp = isp - 1

        if iEmbodimentsManager.CurrentEmbodimentIndex !=0:                      # индекс текущего исполнения
            iEmbodimentsManager.SetCurrentEmbodiment(0)                         # сделать текущим исполнение "0-ое"
    else:
        record(obozn, name, var, docCode, isp, whitespacedocCode)

def RecordMarking(tuple_marking, iPart7): # запись обозначения, исполнения и кода документа

    iMarking = tuple_marking[0] # обозначение
    embodiment = tuple_marking[1] # исп.
    whitespaceDocCode = tuple_marking[2] # пробел перед кодом документа
    docCode = tuple_marking[3] # код докумета

    full_marking = iMarking + "$|-$|" + embodiment + "$|$|$|" + whitespaceDocCode + "$|" + docCode # записываемое обозначение

    print(f"Обозначение: \"{iPart7.Marking}\" => \"{iMarking}-{embodiment}{whitespaceDocCode}{docCode}\" изменено!")

    iPart7.Marking = full_marking # установить обозначение

    iPart7.Update() # применить обозначение

def Сhange_property_SP(iPropertyKeeper, iProperty, cell): # изменение св-в раздела спецификации (интерфейс получения/редактирования значения свойств, интерфейс свойства, значение ячейки)

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

#-------------------------------------------------------------------------------

list_Excel = [] # список считанных строк из Excel
list_files = [] # список считаных файлов

DoubleExe() # проверка на уже запущеное приложени

Check_update() # проверить обновление приложение

KompasAPI() # подключение API компаса

SystemPath() # определяем папку системных файлов

Сhecking_open_files() # проверка открытых файлов

Connecting_to_Excel() # подключение EXcel и изменение св-в

Main_assembly() # обработка главной СБ или дет.
