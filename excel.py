import openpyxl
import re
from openpyxl import Workbook as ws
from openpyxl import load_workbook

#Импорт excel файла для этикеток ^-^
workbook = openpyxl.load_workbook('wildberries.xlsx')

#::: Выбираю лист который мне нужен :::
sheet = workbook['Лист подбора']

#___Список для хранения данных из столбца___
column_l = []
M = []

# '''Проходим по всем строкам из столбца L'''
for cell in sheet[
    "G"]:
    #добавляю значение ячейки в список
    if cell.value == "Артикул продавца":
        continue
    column_l.append(cell.value)

# Функция для удаления знаков из названия файла
def remove_special_chars(name):
    # используем регулярное выражение для поиска запрещенных знаков в имени файла

    return re.sub('[\\/:*?"<>|]', '', name)

for file in column_l:
    # Проверяем, что значение не является None
    if file is not None:
        # удаляем запрещенные знаки из имени файла
        file_name = remove_special_chars(file)
        # добавляем новое имя файла в список
        M.append(file_name + '.rtf')
