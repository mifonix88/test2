# -*- coding: utf-8 -*-
import csv
import sys
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
import subprocess

def read_file(file_name):

    
    values = {}
    number_id = {}
    
    
    with open(file_name, encoding='utf-8') as r_file:
        reader = csv.DictReader(r_file, delimiter = ";")
        
        temp = {}
        
        for i in reader:
            art = int(i['Артикул'])
            name = i['Наименование товара']
            kol = int(i['Количество'])
            id_ = i['Номер отправления']
            
            if values.get(art):
                values[art].append(kol)
            else:
                values[art] = [name, kol]
            
            if temp.get(id_):
                temp[id_].append((art ,kol))
            else:
                temp[id_] = [(art ,kol)]
        
        for i in temp.keys():
            list_ = temp[i]

            if len(list_) > 1: 
                number_id[i] = list_
            elif list_[0][1] > 1: 
                number_id[i] = list_

    return  values, number_id

def path_to_root(name):
    #принемает путь до открываемого файла
    #вернёт путь для сохранения одноимённого файла в корень
    name_file = name[:-4].split('\\')[-1]#имя файла без расширения(-4)
    root = sys.argv[0].split('\\')[:-1]
    root.append(name_file)
    name = '\\'.join(root)
    return name

def save_files(name, val, num):
    wb = Workbook()
    ws = wb.active
    ws.title = "Общий"
    ws.column_dimensions['B'].width = 80
    ws.column_dimensions['A'].width = 20
    
    ws1 = wb.create_sheet("На печать")
    ws1.column_dimensions['B'].width = 80
    ws1.column_dimensions['A'].width = 20

    ws2 = wb.create_sheet("На лазер")
    ws2.column_dimensions['B'].width = 80
    ws2.column_dimensions['A'].width = 20

    key = val.keys()
    key = sorted(key)
    
    temp_in = len(key) + 3
    
    num_val = {}
    arg_ = num.values()
    for i in arg_:
        for j in i:
            k, *v = j
            if num_val.get(k):
                num_val[k].append(v[0])
            else:
                num_val[k] = v


    
    def top(ls, ws):
        for i in ls:
            arg = ws.cell(row=1, column=1+ls.index(i), value = i)
            arg.font = Font(size= 12, bold=True)
            arg.alignment = Alignment(horizontal='center')
    ls = ('Арт.', 'Наименование', 'Всего:', 'По 1шт', 'Сборные')
    top(ls,ws)
    top(ls[:3],ws1)
    top(ls[:3],ws2)

    nums = 2
    nums2 = 2
    

    str_ = open_cofig(2)
    list_to_print = [int(i) for i in str_.split(',')]#список для печати
    
    str_ = open_cofig(3)
    list_to_laser = [int(i) for i in str_.split(',')]

    for index, art in enumerate(key):

        index = index + 2

        arg1 = ws.cell(row=index, column=1, value = art)
        arg1.alignment = Alignment(horizontal='center')
        
        arg2 = ws.cell(row=index, column=2, value = val[art][0])
        arg2.alignment = Alignment(horizontal='left')

        arg3 = ws.cell(row=index, column=3, value = sum(val[art][1:]))
        arg3.alignment = Alignment(horizontal='center')
        arg3.font = Font(bold=True)
        
        if art in list_to_print:#если есть в этом списке дублируем на страницу
            
            arg11 = ws1.cell(row=nums, column=1, value = art)
            arg11.alignment = Alignment(horizontal='center')
            
            arg12 = ws1.cell(row=nums, column=2, value = val[art][0])
            arg12.alignment = Alignment(horizontal='left')

            arg13 = ws1.cell(row=nums, column=3, value = sum(val[art][1:]))
            arg13.alignment = Alignment(horizontal='center')
            arg13.font = Font(bold=True)
            
            nums +=1

        elif art in list_to_laser:#если есть в этом списке дублируем на страницу
            
            arg11 = ws2.cell(row=nums2, column=1, value = art)
            arg11.alignment = Alignment(horizontal='center')
            
            arg12 = ws2.cell(row=nums2, column=2, value = val[art][0])
            arg12.alignment = Alignment(horizontal='left')

            arg13 = ws2.cell(row=nums2, column=3, value = sum(val[art][1:]))
            arg13.alignment = Alignment(horizontal='center')
            arg13.font = Font(bold=True)
            
            nums2 +=1
        
        sb = []
        if num_val.get(art):
            sb = sb + num_val[art]
        sb2 = str(sb) if sb else ''

        
        arg5 = ws.cell(row=index, column=5, value = sb2)
        arg5.alignment = Alignment(horizontal='center')
        
    
        su = sum(val[art][1:]) - sum(sb)
        su = su if su else ''
        arg4 = ws.cell(row=index, column=4, value = su)
        arg4.alignment = Alignment(horizontal='center')

    _ = ws.cell(row=temp_in, column=2, value = f"Список сборных заказов (Всего: {len(num)})")
    _.font = Font(size= 14, bold=True)
    _.alignment = Alignment(horizontal='center')

    temp_in +=1
    
    ls = ('Отправления', 'Наименование', 'Кол-во')
    for i in ls:
        _ = ws.cell(row=temp_in, column=1+ls.index(i), value = i)
        _.font = Font(size= 12, bold=True)
        _.alignment = Alignment(horizontal='center')


    for i in num:
        
        if len(num[i]) > 1:
            temp_in +=2
        else:
            temp_in +=1
        
        a = ws.cell(row=temp_in, column=1, value = i)
        a.alignment = Alignment(horizontal='center')

        for j in num[i]:
            art , kol = j
            b = ws.cell(row=temp_in, column=2, value = val[art][0])
            b.alignment = Alignment(horizontal='left')
            
            c = ws.cell(row=temp_in, column=3, value = kol)
            c.alignment = Alignment(horizontal='center')
            
            if len(num[i]) > 1:
                temp_in +=1
            
            

    try:
        os.remove(f'{name}.xlsx')
    except(FileNotFoundError):pass

    wb.save(f'{name}.xlsx')
 
def open_cofig(num):
    path_to_konfig = path_to_root('config....')#будет искать в корне
    
    with open(path_to_konfig) as _file:
        lins = _file.readlines()

    return lins[num-1].rstrip('\n')
    


def open_excel(name):
    #Открывает файл excel по указаному адресу
    lins = open_cofig(1)#ищем адрес в файле в первой строке
    subprocess.Popen([lins, f'{name}.xlsx'])

def main():
    try:
        name = sys.argv[1]#получаем путь к открываемому файлу
        values, number_id = read_file(name)#получаем значения из файла
        name = path_to_root(name)#получаем адрес в корень без 4х последних символов(расширение)
        save_files(name, values, number_id)#записываем и сохраняем
        open_excel(name)#открываем с помощью EXCEL
    except(IndexError):pass



if __name__ == '__main__':
    main()



