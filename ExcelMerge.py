# -*- coding: utf-8 -*-
#from __future__ import unicode_literals  # Для того чтобы не использовать u''
import pyexcel
import pyexcel.ext.xls
import os
import glob

curdir = os.path.dirname(__file__)
os.chdir(curdir)
infile = os.path.join(curdir,'out','7.1.xls')
outfile = os.path.join(curdir,'out','32out.xls')
X_MIN = 3 # Номер столбца для начала обработки
X_MAX = 40 # Номер столбца для конца обработки
Y_MIN = 7 # Номер начальной строки
Y_MAX = 21 # Номер конечной строки

#if not os.path.exists('out'):
#    os.mkdir('out')

print 'Текущая папка:'.decode('utf-8'),
print curdir.decode('cp1251')
def test():
    sheet_out = pyexcel.get_sheet(file_name=infile)
    files = glob.glob('*.xls*')
    sheets = []
    for fn in files:
        rec = {}
        rec['filename'] = fn
        print 'Открытие файла:'.decode('utf-8'),
        print(fn.decode('cp1251'))
        rec['sheet'] = pyexcel.get_sheet(file_name=fn)
        sheets.append(rec)
        
    for y in range(Y_MIN-1,Y_MAX): #5,33 # Обрабатывается с 6 по 33 строку
        for x in range(X_MIN-1,X_MAX): #2,9 # Обрабатывается с 3 по 9 колонки
            try:
                item_sum = 0
                print '[{}:{}]'.format(x,y),
                for sheet in sheets:
                    item_val = sheet['sheet'][y,x]
                    if type(item_val) != float:
                        item_val = 0
                    item_sum += item_val
                    print item_val,
                print 'SUM:', item_sum
                sheet_out[y,x] = item_sum
            except Exception as e:
                print '!!!Ошибка обработки'.decode('utf-8'), e
    sheet_out.save_as(outfile)
    return
    
            
if __name__=='__main__':
    test()
    raw_input()
    #print('Ok')
