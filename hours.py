{
 "cells": [
  {
   "cell_type": "markdown",
   "id": "3d0bea65",
   "metadata": {},
   "source": [
    "# Отчёт по часам"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 141,
   "id": "6de3ac18",
   "metadata": {},
   "outputs": [],
   "source": [
    "# Импортим пандас и загружаем файлы через переменные\n",
    "import pandas as pd\n",
    "path = ()\n",
    "week = '31'\n",
    "filename = '31_1'\n",
    "df = pd.read_excel('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' + week + '/'+ filename + '.xlsx', skiprows = 3)\n",
    "df_voc = pd.read_excel('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' + week + '/'+ 'vocab.xlsx')\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 142,
   "id": "4efdc618",
   "metadata": {
    "scrolled": true
   },
   "outputs": [],
   "source": [
    "# Филтруем лишнее и делаем лефт джоин\n",
    "df = df[['Магазин', \n",
    "         'Должность',\n",
    "         'Факт. отработ. время без переработок(час)',\n",
    "         'Переработки(час)',\n",
    "         'Факт. время внеш. персонала(час)',\n",
    "         'Центр финансовой отчетности', \n",
    "         'Код SAP']]\n",
    "\n",
    "df = pd.merge(df, df_voc, left_on = 'Магазин', right_on = 'Магазин', how = 'left')\n",
    "df = df.drop(columns = ['Магазин'], axis = 1)\n",
    "df.rename(columns = {'Unnamed: 1':'РЦ'}, inplace = True)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 143,
   "id": "3826cda7",
   "metadata": {
    "scrolled": false
   },
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "['Должность', 'Факт. отработ. время без переработок(час)', 'Переработки(час)', 'Факт. время внеш. персонала(час)', 'Центр финансовой отчетности', 'Код SAP', 'РЦ']\n"
     ]
    }
   ],
   "source": [
    "print(list(df.columns))"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 144,
   "id": "ffaeb52c",
   "metadata": {
    "scrolled": false
   },
   "outputs": [],
   "source": [
    "# Небольшой решейп\n",
    "# print(list(df.columns))\n",
    "\n",
    "df = df.iloc[:, [6, 0, 1, 2, 3, 4, 5]]"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 145,
   "id": "04430816",
   "metadata": {
    "scrolled": false
   },
   "outputs": [],
   "source": [
    "# print(df_voc) # Смотрим если нужно, что в справочнике"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 146,
   "id": "0213cb10",
   "metadata": {
    "scrolled": true
   },
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Index(['РЦ', 'Должность', 'Факт. отработ. время без переработок(час)',\n",
      "       'Переработки(час)', 'Факт. время внеш. персонала(час)',\n",
      "       'Центр финансовой отчетности', 'Код SAP'],\n",
      "      dtype='object')\n"
     ]
    }
   ],
   "source": [
    "print(df.columns)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 147,
   "id": "8b4dbd35",
   "metadata": {},
   "outputs": [],
   "source": [
    "df_sof = df[df['РЦ'] == 'РЦ Софьино']\n",
    "df_frov = df[df['РЦ'] == 'РЦ ФРОВ']\n",
    "df_nn = df[df['РЦ'] == 'РЦ НН']\n",
    "df_adg = df[df['РЦ'] == 'РЦ Адыгея']\n",
    "df_vor = df[df['РЦ'] == 'РЦ Воронеж']\n",
    "df_slk = df[df['РЦ'] == 'РЦ СЛК']\n",
    "df_nor = df[df['РЦ'] == 'РЦ Северный']\n",
    "df_spb = df[df['РЦ'] == 'РЦ СПб']\n",
    "df_ekb = df[df['РЦ'] == 'РЦ Косулино']"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 148,
   "id": "aea17206",
   "metadata": {
    "scrolled": true
   },
   "outputs": [],
   "source": [
    "# Создаём пачку врайтеров, чтобы записать разные файлы\n",
    "\n",
    "writer_sof = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_Софьино' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')\n",
    "writer_frov = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_ФРОВ' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')\n",
    "writer_nn = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_НН' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')\n",
    "writer_adg = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_Адыгея' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')\n",
    "writer_vor = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_Воронеж' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')\n",
    "writer_slk = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_СЛК' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')\n",
    "writer_nor = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_Север' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')\n",
    "writer_spb = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_СПб' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')\n",
    "writer_ekb = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_ЕКб' + '_' + week + 'W' + '.xlsx'), engine='xlsxwriter')\n",
    "\n",
    "writer = pd.ExcelWriter(('C:/Users/Alexander.Persidskiy/Desktop/Python/DF/weeks/' \n",
    "                             + week + '/' + 'output/РЦ_ALL.xlsx'), engine='xlsxwriter')\n",
    "\n",
    "# Записываем листы в каждый файл\n",
    "df_sof.to_excel(writer_sof, sheet_name='РЦ Софьино')\n",
    "df_frov.to_excel(writer_frov, sheet_name='РЦ ФРОВ')\n",
    "df_nn.to_excel(writer_nn, sheet_name='РЦ НН')\n",
    "df_adg.to_excel(writer_adg, sheet_name='РЦ Адыгея')\n",
    "df_vor.to_excel(writer_vor, sheet_name='РЦ Воронеж')\n",
    "df_slk.to_excel(writer_slk, sheet_name='РЦ СЛК')\n",
    "df_nor.to_excel(writer_nor, sheet_name='РЦ Северный')\n",
    "df_spb.to_excel(writer_spb, sheet_name='РЦ СПб')\n",
    "df_ekb.to_excel(writer_ekb, sheet_name='РЦ Косулино')\n",
    "df.to_excel(writer, sheet_name = 'All')\n",
    "\n",
    "# Записываем файл\n",
    "writer_sof.save()\n",
    "writer_frov.save()\n",
    "writer_nn.save()\n",
    "writer_adg.save()\n",
    "writer_vor.save()\n",
    "writer_slk.save()\n",
    "writer_nor.save()\n",
    "writer_spb.save()\n",
    "writer_ekb.save()\n",
    "writer.save()\n",
    "\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 149,
   "id": "ac14040a",
   "metadata": {},
   "outputs": [
    {
     "name": "stderr",
     "output_type": "stream",
     "text": [
      "C:\\Users\\Alexander.Persidskiy\\Anaconda3\\lib\\site-packages\\xlsxwriter\\workbook.py:339: UserWarning: Calling close() on already closed file.\n",
      "  warn(\"Calling close() on already closed file.\")\n"
     ]
    }
   ],
   "source": [
    "writer_sof.close()\n",
    "writer_frov.close()\n",
    "writer_nn.close()\n",
    "writer_adg.close()\n",
    "writer_vor.close()\n",
    "writer_slk.close()\n",
    "writer_nor.close()\n",
    "writer_spb.close()\n",
    "writer_ekb.close()\n",
    "writer.close()"
   ]
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.9.12"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
