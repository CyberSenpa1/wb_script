from excel import M, get_shelf
import os
import win32com.client as win32
import pythoncom
import time

# Получаем текущую директорию, где находится скрипт
script_directory = os.path.dirname(os.path.abspath(__file__))

# Строим путь к папке "Артикулы"
artikuly_folder = os.path.join(script_directory, 'Артикулы')

# Инициализируем Word
word = win32.gencache.EnsureDispatch('Word.Application')

# Скрываем окно Word
word.Visible = False

# Отключаем все диалоговые окна
word.DisplayAlerts = False

# Список для хранения названий файлов, которые не найдены
not_found = []

# Проходим по всем именам файлов в списке M
for filename in M:
    # Создаем полный путь к файлу
    file_path = os.path.join(artikuly_folder, filename)
    
    # Проверяем наличие файла
    if not os.path.isfile(file_path):
        # Если файл не найден, добавляем его название в список not_found
        not_found.append(filename)
        continue
    
    # Открываем файл Word
    try:
        doc = word.Documents.Open(file_path)
    except Exception as e:
        print('Ошибка при открытии файла', e)
        continue
    
    # Получаем артикул из имени файла
    articul = filename[:-4]  # Убираем расширение .rtf
    
    # Отладочная информация
    print(f"Обрабатываемый артикул: {articul}")
    
    # Получаем стеллаж по артикулу
    shelf = get_shelf(articul)
    
    # Отладочная информация
    print(f"Найденный стеллаж: {shelf}")
    
    # Добавляем информацию о стеллаже в конец документа
    try:
        # Перемещаем курсор в конец документа
        end_range = doc.Content
        end_range.Collapse(Direction=0)  # 0 означает конец документа
        
        # Добавляем текст "Стеллаж: <номер стеллажа>"
        end_range.InsertAfter("\n")
        end_range.InsertAfter(f"Стеллаж: {shelf}")
        
        
        # Настраиваем шрифт
        end_range.Font.Size = 5  # 5 размер пунктов
        end_range.Font.Name = "Times New Roman"  # Укажите нужный шрифт
        
        # Выравниваем текст по левому краю
        end_range.ParagraphFormat.Alignment = 0  # 0 = выравнивание по левому краю
    except Exception as e:
        print(f'Ошибка при добавлении информации о стеллаже в файл {filename}: {e}')
    
    # Печатаем содержимое файла
    try:
        # Печать без диалоговых окон
        doc.PrintOut(Background=False)
    except Exception as e:
        print(f'Ошибка при печати файла {filename}: {e}')
    
    # Закрываем файл Word без сохранения изменений
    try:
        doc.Close(win32.constants.wdDoNotSaveChanges)  # Не сохраняем изменения
    except Exception as e:
        print('Ошибка при закрытии файла', e)
    
    time.sleep(1)
    

# Закрываем экземпляр Word.Application
word.Quit()

# Закрываем все объекты COM
pythoncom.CoUninitialize()

# Выводим список ненайденных файлов
if not_found:
    print('Следующие файлы не найдены:')
    for filename in not_found:
        print(filename)
else:
    print('Все этикетки успешно распечатаны')