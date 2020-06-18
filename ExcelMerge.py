"""
    Объединение файлов Excel и OpenOffice путем суммирования значений во всех листах с выгрузкой результата в файл RESULT.xls.
    Применяется для получения сводных или итоговых данных из множества однотипных файлов (с одинаковыми листами и расположением данных)
    Форматирование в результирующем файле при этом теряется.
    
    Использование:
        Помещаем файл ExcelMerge.py в каталог с файлами *.xls, *.xlsx, *.ods и запускаем его
        В результате в текущем каталоге создастся файл RESULT.xls и автоматически откроется.
    
    Примечания:
        Во всех файлах должны совпадать имена Листов
        По умолчанию обработка ведется начиная со второго столбца (B) второй строки (Можно изменить в переменных 
            X_MIN, X_MAX, Y_MIN, Y_MAX)

    Установка:
        Для использования данного скрипта, необходим Python3 и библиотека pyexcel
        Python3 можно скачать по ссылке: https://www.python.org/downloads/ затем установить
        Для установки расширения pyexcel выполнить в командной строке команду: 
            pip install pyexcel pyexcel-xls pyexcel-xlsx pyexcel-ods
            
    Данный скрипт был протестирован на Python 3.6, pyexcel==0.6.2
"""

import os, sys
import pyexcel
import glob

curdir = os.path.abspath(os.path.dirname(__file__)) # Текущий каталог, где лежит этот файл

# === Настройки ===
WORKDIR = os.path.join(curdir, '!DATA') # Рабочий каталог, он же текущий по умолчанию

INFILE = None # Начальный файл. Может быть None, тогда начальный файл считается первый по списку # os.path.join(curdir,'out','infile.xlsx')
OUTFILE = 'RESULT.xls' # Имя исходящего результирующего файла XLS (может находиться в другой папке) # os.path.join(curdir,'out','out.xls')

X_MIN = 2 # Номер столбца для начала обработки
X_MAX = 100000 # Максимальный Номер столбца для конца обработки
Y_MIN = 2 # Номер начальной строки
Y_MAX = 100000 # Максимальный Номер конечной строки

ISSAVE_FILE = False  # Сохранять файл
ISOPEN_FILE = True  # Открыть результирующий файл после обработки?
ISNOTCLOSE = True   # После обработки файла не закрывать окно скрипта

CELL_ANALIZE = 'E2'
SHEET_ANALIZE = 'МТБ'

def parse_float(x):
    if x == '':
        res = 0
    else:
        try:
            res = float(x)
        except:
            res = None
    return res


def cell2idx(cellname):
    """ Возвращает индекс ячейки по x и по y в виде кортежа (x, y) 
        Индексы начинаются с 0
        Например:
            cell2idx('A1') = (0, 0)
            cell2idx('C10') = (2, 9)
            cell2idx('BBB58') = (1405, 57)
    """
    import string
    col = ''
    col_idx = 0
    row = ''
    row_idx = 0
    i = 0
    for l in cellname.upper()[::-1]:
        num = string.ascii_uppercase.find(l)
        if num<0:
            row += l
        else:
            col += l
            col_idx += (26**i)*(num+1)
            i += 1
    col_idx -= 1
    row_idx = int(row[::-1])-1
    #col = col[::-1]
    #row = row[::-1]
    #print(i, l, num, col_idx)
    #print(cellname)
    #print(col, row)
    return (col_idx, row_idx)
  

def merge_all_sheets():
    """ Складывает все значения в файлах Excel в рабочей директории по всем листам и записывает результат в RESULT.xls
          Названия листов во всех файлах должны совпадать, так же как и диапазоны строк и столбцов """
    err_files = set()
    err_all = 0
    files_count = 0
    book_res = None  # Рабочий массив данных
    if INFILE:  # В рабочий массив загружаем данные из начального файла, если он задан
        book_res = pyexcel.get_book_dict(file_name=INFILE)

    files = glob.glob('*.xls*')
    files.extend(glob.glob('*.ods'))
    if files:
        for fn in files:
            if not fn.startswith('~') and fn != OUTFILE and fn != INFILE:
                if not CELL_ANALIZE:
                    print(fn)
                files_count += 1
                if book_res is None:
                    book_res = pyexcel.get_book_dict(file_name=fn)
                    
                    if CELL_ANALIZE:
                        x_idx, y_idx = cell2idx(CELL_ANALIZE)
                        for sheet in book_res:
                            for y, row in enumerate(book_res[sheet]):
                                for x, cell in enumerate(book_res[sheet][y]):
                                    if y>=Y_MIN-1 and y<=Y_MAX-1 and x>=X_MIN-1 and x<=X_MAX-1:
                                        if x == x_idx and y == y_idx:
                                            if SHEET_ANALIZE==sheet or SHEET_ANALIZE=='':
                                                print('Значение: {:<5}  {} ({},{}) Лист="{}" Файл="{}"'.format(book_res[sheet][y][x], CELL_ANALIZE, x, y, sheet, fn))

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
                                            if CELL_ANALIZE:
                                                x_idx, y_idx = cell2idx(CELL_ANALIZE)
                                                if x == x_idx and y == y_idx:
                                                    if SHEET_ANALIZE==sheet or SHEET_ANALIZE=='':
                                                        print('Значение: {:<5}  {} ({},{}) Лист="{}" Файл="{}"'.format(book_cur[sheet][y][x], CELL_ANALIZE, x, y, sheet, fn))
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

        if ISSAVE_FILE:
            pyexcel.save_book_as(bookdict=book_res, dest_file_name=OUTFILE)

            if ISOPEN_FILE:
                os.startfile(OUTFILE)
    else:
        print('В папке {} файлов не найдено'.format(WORKDIR))
        if sys.platform == 'win32': # Если Windows открывать пустую папку Explorer'ом
            os.system('explorer {}'.format(WORKDIR))

def merge_first_sheet():
    """ Объединяет значения в файлах Excel в рабочей директории и записывает результат в RESULT.xls
          Берутся только первые листы
    """
    err_files = set()
    err_all = 0
    files_count = 0
    sheet_res = None  # Рабочий массив данных
    if INFILE:  # В рабочий массив загружаем данные из начального файла, если он задан
        sheet_res = pyexcel.get_book_dict(file_name=INFILE)
        #sheet_name =

    files = glob.glob('*.xls*')
    if files:
        sheets = []
        for fn in files:
            if not fn.startswith('~') and fn != OUTFILE and fn != INFILE:
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

        pyexcel.save_as(array=sheet_res, dest_file_name=OUTFILE)

        if ISOPEN_FILE:
            os.startfile(OUTFILE)
    else:
        print('В папке {} файлов не найдено'.format(WORKDIR))
        if sys.platform == 'win32': # Если Windows открывать пустую папку Explorer'ом
            os.system('explorer {}'.format(WORKDIR))


def main():
    print('Версия Python:', sys.version)
    print('Текущая папка:', curdir)  # .decode('utf-8'),
    print('Рабочая папка:', WORKDIR)  # .decode('utf-8'),
    if WORKDIR:
        if not os.path.exists(WORKDIR):
            os.makedirs(WORKDIR, exist_ok=True) #
            print('Создан каталог: {}'.format(WORKDIR))
        os.chdir(WORKDIR) # Работа в текущем каталоге
    print()
    # Слить все файлы в рабочей директории со всеми листами
    merge_all_sheets()
    # Слить все файлы в рабочей директории только по первым листам
    #merge_first_sheet()
    if ISNOTCLOSE:
        input('Для продолжения нажмите любую клавишу...')


if __name__=='__main__':
    main()

