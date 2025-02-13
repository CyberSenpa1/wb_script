from excel import M
import os
import win32com.client as win32
import pythoncom
import time

word = win32.Dispatch('Word.Application')

# список для хранения названий файлов, которые не найдены
not_found = []

# проходим по всем именам файлов в списке L
for filename in M:
    # создаем полный путь к файлу
    file_path = os.path.join('C:\\Users\\Voro_\\OneDrive\\Desktop\\печать\\Программа\\Артикулы', filename)
    
    # проверяем наличие файла
    if not os.path.isfile(file_path):
        # если файл не найден, добавляем его название в список not_found
        not_found.append(filename)
        continue
    
    # открываем файл Word
    try:
        doc = word.Documents.Open(file_path)
    except Exception as e:
        print('Ошибка пока отрывался файл', e)
        continue
    
    # печатаем содержимое файла
    try:
        doc.PrintOut()
    except Exception as e:
        print(f'Ошибка при печати файла {filename}: {e}')
    
    # закрываем файл Word
    try:
        doc.Close()

    except Exception as e:
        print('Ошибка пока файл закрывался', e)
    
    time.sleep(1)
    

# закрываем экземпляр Word.Application
word.Quit()

# закрываем все объекты COM
pythoncom.CoUninitialize()

# выводим список ненайденных файлов
if not_found:
    print('Следующие файлы не найдены:')
    for filename in not_found:
        print(filename)
else:
    print('Все этикетки успешно распечатаны')