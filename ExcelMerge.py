# -*- coding: utf-8 -*-
#from __future__ import unicode_literals  # Для того чтобы не использовать u''
import os, sys
print(sys.version)
import pyexcel
#import pyexcel.ext.xls # Deprecated usage since v0.2.2! Explicit import is no longer required. pyexcel.ext.xls is auto imported.
import glob

curdir = os.path.dirname(__file__) # Текущий каталог, где лежит программа
print(curdir)
print(__file__)
workdir = curdir # Рабочий каталог, он же текущий по умолчанию

# === Настройки ===

#workdir = r'e:\NEW\2016-07-18\nika\MY'  # Рабочий каталог
#if workdir:
os.chdir(workdir) # Работа в текущем каталоге

infile = None # Начальный файл. Может быть None, тогда начальный файл считается первый по списку # os.path.join(curdir,'out','infile.xlsx')
outfile = 'RESULT.xls' # Имя исходящего результирующего файла XLS (пожжет быть в другой папке) # outfile = os.path.join(curdir,'out','out.xls')

X_MIN = 2 # Номер столбца для начала обработки
X_MAX = 100000 # Максимальный Номер столбца для конца обработки
Y_MIN = 2 # Номер начальной строки
Y_MAX = 100000 # Максимальный Номер конечной строки

isopen_file = True  # Открыть результирующий файл после обработки?
isnotclose = True  # После обработки файла не закрывать окно скрипта

# ПРИМЕЧАНИЯ
# Во всех файлах должны совпадать имена Листов

#if not os.path.exists('out'):
#    os.mkdir('out')


def parse_float(x):
    if x == '':
        res = 0
    else:
        try:
            res = float(x)
        except:
            res = None
    return res


def merge_all_sheets():
    """ Складывает все значения в файлах Excel в рабочей директории по всем листам и записывает результат в RESULT.xls
          Названия листов во всех файлах должны совпадать, так же как и диапазоны строк и столбцов """
    err_files = set()
    err_all = 0
    files_count = 0
    book_res = None  # Рабочий массив данных
    if infile:  # В рабочий массив загружаем данные из начального файла, если он задан
        book_res = pyexcel.get_book_dict(file_name=infile)

    files = glob.glob('*.xls*')
    #sheets = []
    for fn in files:
        if not fn.startswith('~') and fn != outfile and fn != infile:
            print(fn)
            files_count += 1
            if book_res is None:
                book_res = pyexcel.get_book_dict(file_name=fn)
            else:
                book_cur = pyexcel.get_book_dict(file_name=fn)
                for sheet in book_res:
                    #print(sheet)
                    for y, row in enumerate(book_res[sheet]):
                        #print(y, row)
                        for x, cell in enumerate(book_res[sheet][y]):
                            if y>=Y_MIN-1 and y<=Y_MAX-1 and x>=X_MIN-1 and x<=X_MAX-1:  # Ограничиваем обработку только определенным диапазоном
                                try:
                                    cell1 = parse_float(book_res[sheet][y][x])
                                    cell2 = parse_float(book_cur[sheet][y][x])
                                    if cell1 is not None and cell2 is not None:
                                        book_res[sheet][y][x] = cell1 + cell2
                                        #print(x, book_res[sheet][y][x], type(book_res[sheet][y][x]))
                                except:
                                    print('Ошибка! Файл: {} Лист: {} Строка: {} Колонка: {}'.format(fn, sheet, y, x))
                                    err_all += 1
                                    err_files.add(fn)
                    #print(book_res[sheet])
    print('\nОбработка завершена!\n  Всего обработано файлов: {}'.format(files_count))
    if err_all > 0:
        print('  Всего ошибок: {}'.format(err_all))
        print('  Ошибки в файлах: {}'.format(', '.join(err_files)))
    else:
        print('  Ошибок нет')
    pyexcel.save_book_as(bookdict=book_res, dest_file_name=outfile)

    if isopen_file:
        import os
        os.startfile(outfile)


def merge_first_sheet():
    """ Объединяет значения в файлах Excel в рабочей директории и записывает результат в RESULT.xls
          Берутся только первые листы
    """
    err_files = set()
    err_all = 0
    files_count = 0
    sheet_res = None  # Рабочий массив данных
    if infile:  # В рабочий массив загружаем данные из начального файла, если он задан
        sheet_res = pyexcel.get_book_dict(file_name=infile)
        #sheet_name =

    files = glob.glob('*.xls*')
    sheets = []
    for fn in files:
        if not fn.startswith('~') and fn != outfile and fn != infile:
            print(fn)
            files_count += 1
            if sheet_res is None:
                #sheet_res = pyexcel.get_sheet(file_name=fn)
                sheet_res = pyexcel.get_array(file_name=fn)
            else:
                sheet_cur = pyexcel.get_array(file_name=fn)
                # Определение названия листа (берется первый)
                sheet_name = list(pyexcel.get_book_dict(file_name=fn).keys())[0]
                #print(sheet_name)
                #print(sheet_res)
                for y, row in enumerate(sheet_res):
                    #print(y, row)
                    for x, cell in enumerate(sheet_res[y]):
                        if y>=Y_MIN-1 and y<=Y_MAX-1 and x>=X_MIN-1 and x<=X_MAX-1:  # Ограничиваем обработку только определенным диапазоном
                            try:
                                cell1 = parse_float(sheet_res[y][x])
                                cell2 = parse_float(sheet_cur[y][x])
                                if cell1 is not None and cell2 is not None:
                                    sheet_res[y][x] = cell1 + cell2
                                    #print(x, sheet_res[y][x], type(sheet_res[y][x]))
                            except:
                                print('Ошибка! Файл: {} Лист: {} Строка: {} Колонка: {}'.format(fn, sheet_name, y, x))
                                err_all += 1
                                err_files.add(fn)
                #print(book_res[sheet])
    print('\nОбработка завершена!\n  Всего обработано файлов: {}'.format(files_count))
    if err_all > 0:
        print('  Всего ошибок: {}'.format(err_all))
        print('  Ошибки в файлах: {}'.format(', '.join(err_files)))
    else:
        print('  Ошибок нет')

    pyexcel.save_as(array=sheet_res, dest_file_name=outfile)

    if isopen_file:
        import os
        os.startfile(outfile)


def main():
    print('Версия Python:', sys.version)
    print('Текущая папка:', curdir)  # .decode('utf-8'),
    print('Рабочая папка:', workdir)  # .decode('utf-8'),
    print()
    # Слить все файлы в рабочей директории со всеми листами
    merge_all_sheets()
    # Слить все файлы в рабочей директории только по первым листам
    #merge_first_sheet()


if __name__=='__main__':
    main()
    if isnotclose:
        input()

