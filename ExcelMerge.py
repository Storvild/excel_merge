# -*- coding: utf-8 -*-
#from __future__ import unicode_literals  # Для того чтобы не использовать u''
import pyexcel
#import pyexcel.ext.xls # Deprecated usage since v0.2.2! Explicit import is no longer required. pyexcel.ext.xls is auto imported.
import os
import glob

curdir = os.path.dirname(__file__) # Текущий каталог, где лежит программа
os.chdir(curdir) # Работа в текущем каталоге

infile = None # Начальный файл. Может быть None, тогда начальный файл считается первый по списку # os.path.join(curdir,'out','infile.xlsx')
outfile = 'RESULT.xls' # Имя исходящего результирующего файла XLS (пожжет быть в другой папке) # outfile = os.path.join(curdir,'out','out.xls')

X_MIN = 2 # Номер столбца для начала обработки
X_MAX = 100000 # Максимальный Номер столбца для конца обработки
Y_MIN = 2 # Номер начальной строки
Y_MAX = 100000 # Максимальный Номер конечной строки

open_file = True  # Открыть результирующий файл после обработки?


#if not os.path.exists('out'):
#    os.mkdir('out')

print ('Текущая папка:') #.decode('utf-8'),
print (curdir) #.decode('cp1251')
def test():
    sheet_out = pyexcel.get_sheet(file_name=infile)
    files = glob.glob('*.xls*')
    sheets = []
    for fn in files:
        rec = {}
        rec['filename'] = fn
        print ('Открытие файла:') #.decode('utf-8'),
        print(fn) #.decode('cp1251')
        rec['sheet'] = pyexcel.get_sheet(file_name=fn)
        sheets.append(rec)
    
    for y in range(Y_MIN-1,Y_MAX): #7-1, 21 # Обрабатывается с 6 по 20 строку
        for x in range(X_MIN-1,X_MAX): #3-1, 40 # Обрабатывается с 2 по 39 колонки
            #try:
                item_sum = 0
                print ('[{}:{}]'.format(x,y)) #,
                
                for sheet in sheets:
                    #print(sheet)
                    #return
                    print(x, y, sheet['filename'])
                    
                    item_val = sheet['sheet'][y,x]
                    if type(item_val) != float:
                        item_val = 0
                    item_sum += item_val
                    print (item_val) #,
                
                print ('SUM: '+str(item_sum))
                sheet_out[y,x] = item_sum
                
            #except Exception as e:
                #print ('!!!Ошибка обработки '+str(e)) #.decode('utf-8'), e
            #    pass
    sheet_out.save_as(outfile)
    return

def test2():
    """ Чтение из файла как классов pyexcel"""
    filename = 'Файл 002.xlsx'
    book = pyexcel.get_book(file_name=filename)
    print(book)
    for sheet in book:
        print(sheet)
        for row in sheet:
            print(row)
            for cell in row:
                print(cell)


def test3():
    """ Чтение данных из файла как OrderedDict """
    filename = 'Файл 002.xlsx'
    book = pyexcel.get_book_dict(file_name=filename)
    print(book)

    for sheet in book:
        print(sheet)
        for row in book[sheet]:
            print(row)
            for cell in row:
                print(cell)


def parse_float(x):
    if x == '':
        res = 0
    else:
        try:
            res = float(x)
        except:
            res = None
    return res


def test5():
    filename = 'Файл 002.xlsx'
    #book = pyexcel.get_book(file_name=filename)
    #book = pyexcel.get_book_dict(file_name=filename)
    book_res = None
    book1 = pyexcel.get_book_dict(file_name='001.xlsx')
    book2 = pyexcel.get_book_dict(file_name='Файл 002.xlsx')
    book_res = None

    #sheet = pyexcel.get_sheet(file_name=filename)
    #print(book)
    if book_res is None:
        book_res = book1 # Первую книгу не обрабатываем а сразу её в результат
    #else:
    #   Обработка остальных
    book_cur = book2
    for sheet in book_res:
        print(sheet)
        print(book1[sheet])
        print(book2[sheet])
        for y, row in enumerate(book_res[sheet]):
            #print(y, row)
            for x, cell in enumerate(book_res[sheet][y]):
                if y>=Y_MIN-1 and y<=Y_MAX-1 and x>=X_MIN-1 and x<=X_MAX-1:  # Ограничиваем обработку только определенным диапазоном
                    cell1 = parse_float(book_res[sheet][y][x])
                    cell2 = parse_float(book_cur[sheet][y][x])
                    if cell1 is not None and cell2 is not None:
                        book_res[sheet][y][x] = cell1 + cell2
                        print(x, book_res[sheet][y][x], type(book_res[sheet][y][x]))
        print(book_res[sheet])

    pyexcel.save_book_as(bookdict=book_res, dest_file_name='out.xls')
    pass

def test6():
    #filename = 'Файл 002.xlsx'
    #book = pyexcel.get_book(file_name=filename)
    #book = pyexcel.get_book_dict(file_name=filename)
    book_res = None
    #book1 = pyexcel.get_book_dict(file_name='001.xlsx')
    #book2 = pyexcel.get_book_dict(file_name='Файл 002.xlsx')
    book_res = None
    if infile:
        book_res = pyexcel.get_book_dict(file_name=infile)

    files = glob.glob('*.xls*')
    sheets = []
    for fn in files:
        if not fn.startswith('~') and fn != outfile and fn != infile:
            print(fn)
            if book_res is None:
                book_res = pyexcel.get_book_dict(file_name=fn)
            else:
                book_cur = pyexcel.get_book_dict(file_name=fn)
                for sheet in book_res:
                    print(sheet)
                    #print(book1[sheet])
                    #print(book2[sheet])
                    for y, row in enumerate(book_res[sheet]):
                        #print(y, row)
                        for x, cell in enumerate(book_res[sheet][y]):
                            if y>=Y_MIN-1 and y<=Y_MAX-1 and x>=X_MIN-1 and x<=X_MAX-1:  # Ограничиваем обработку только определенным диапазоном
                                cell1 = parse_float(book_res[sheet][y][x])
                                cell2 = parse_float(book_cur[sheet][y][x])
                                if cell1 is not None and cell2 is not None:
                                    book_res[sheet][y][x] = cell1 + cell2
                                    print(x, book_res[sheet][y][x], type(book_res[sheet][y][x]))
                    print(book_res[sheet])

    pyexcel.save_book_as(bookdict=book_res, dest_file_name=outfile)





if __name__=='__main__':
    #test()
    test6()
    print('ok')
    #input()

