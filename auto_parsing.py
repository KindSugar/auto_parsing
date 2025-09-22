from openpyxl import load_workbook

try:
    wb = load_workbook('avtovoz_cars.xlsx')  # Используем относительный путь
    sheet = wb["Лист1"]
    cells = sheet["A1":"B2155"]

    marka = input('Введите марку автомобиля: ')
    est = []
    for marka_model in cells:
        marka_in_doc, model = marka_model
        if marka_in_doc.value is not None and model.value is not None: # Проверяем, что ячейки не пустые
            if marka.lower() == marka_in_doc.value.lower(): # Сравниваем в нижнем регистре
                est.append(f'{marka_in_doc.value} {model.value}')

    if est:
        print(*est, sep='\n')
    else:
        print('К сожалению, данная марка автомобиля не найдена')

except FileNotFoundError:
    print("Ошибка: Файл avtovoz_cars.xlsx не найден.")
except KeyError:
    print("Ошибка: Лист с именем 'Лист1' не найден в файле Excel.")
except Exception as e:
    print(f"Произошла ошибка: {e}")