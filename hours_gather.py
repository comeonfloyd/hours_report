# Импортим пандас и загружаем файлы через переменные
import os
import pandas as pd
week = '31'
path = ('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' + week + '/FB_LC/')

# Записываем пути и считываем путь
os.chdir(path)
path = os.getcwd()
files = os.listdir(path)
files

# Определяем какие файлы в папке эксель

files_xlsx = [f for f in files if f[-4:] == 'xlsx']
files_xlsx

# Создаем пустой дата фрейм
df = pd.DataFrame()

# Цикл для считывания файлов из папки и записывания в один датафрейм
for f in files_xlsx:
    data = pd.read_excel(f)
    df = df.append(data)
    
# Записываем готовую эксельку чутка её правим и отправляем
path = ('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' + week + '/FB_LC/DF_ALL_' + week + '.xlsx' )
writer = pd.ExcelWriter(path, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Все РЦ')
writer.save()
