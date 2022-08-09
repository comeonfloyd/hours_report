# Импортим пандас и загружаем файлы через переменные
import pandas as pd
path = ()
week = '31'
filename = '31_1'
df = pd.read_excel('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' + week + '/'+ filename + '.xlsx', skiprows = 3)
df_voc = pd.read_excel('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' + week + '/'+ 'vocab.xlsx')

# Филтруем лишнее и делаем лефт джоин
df = df[['Магазин', 
         'Должность',
         'Факт. отработ. время без переработок(час)',
         'Переработки(час)',
         'Факт. время внеш. персонала(час)',
         'Центр финансовой отчетности', 
         'Код SAP']]

df = pd.merge(df, df_voc, left_on = 'Магазин', right_on = 'Магазин', how = 'left')
df = df.drop(columns = ['Магазин'], axis = 1)
df.rename(columns = {'Unnamed: 1':'РЦ'}, inplace = True)

# Небольшой решейп
# print(list(df.columns))

df = df.iloc[:, [6, 0, 1, 2, 3, 4, 5]]

print(df.columns)

df_sof = df[df['РЦ'] == 'РЦ Софьино']
df_frov = df[df['РЦ'] == 'РЦ ФРОВ']
df_nn = df[df['РЦ'] == 'РЦ НН']
df_adg = df[df['РЦ'] == 'РЦ Адыгея']
df_vor = df[df['РЦ'] == 'РЦ Воронеж']
df_slk = df[df['РЦ'] == 'РЦ СЛК']
df_nor = df[df['РЦ'] == 'РЦ Северный']
df_spb = df[df['РЦ'] == 'РЦ СПб']
df_ekb = df[df['РЦ'] == 'РЦ Косулино']

# Создаём пачку врайтеров, чтобы записать разные файлы

writer_sof = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_Софьино' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')
writer_frov = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_ФРОВ' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')
writer_nn = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_НН' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')
writer_adg = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_Адыгея' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')
writer_vor = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_Воронеж' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')
writer_slk = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_СЛК' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')
writer_nor = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_Север' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')
writer_spb = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_СПб' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')
writer_ekb = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_ЕКб' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')

writer = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' 
                             + week + '/' + 'output/РЦ_ALL.xlsx'), engine='xlsxwriter')

# Записываем листы в каждый файл
df_sof.to_excel(writer_sof, sheet_name='РЦ Софьино')
df_frov.to_excel(writer_frov, sheet_name='РЦ ФРОВ')
df_nn.to_excel(writer_nn, sheet_name='РЦ НН')
df_adg.to_excel(writer_adg, sheet_name='РЦ Адыгея')
df_vor.to_excel(writer_vor, sheet_name='РЦ Воронеж')
df_slk.to_excel(writer_slk, sheet_name='РЦ СЛК')
df_nor.to_excel(writer_nor, sheet_name='РЦ Северный')
df_spb.to_excel(writer_spb, sheet_name='РЦ СПб')
df_ekb.to_excel(writer_ekb, sheet_name='РЦ Косулино')
df.to_excel(writer, sheet_name = 'All')

# Записываем файл
writer_sof.save()
writer_frov.save()
writer_nn.save()
writer_adg.save()
writer_vor.save()
writer_slk.save()
writer_nor.save()
writer_spb.save()
writer_ekb.save()
writer.save()
