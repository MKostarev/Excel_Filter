import pandas as pd
import re

# Функция для загрузки данных
def load_data(file_path):
    return pd.read_excel(file_path, header=None)

# Функция для обработки каждой строки
def process_row(row, previous_ut_value, current_receipt):
    row_data = list(row)
    ut_value = None
    code_value = None
    month_year_value = None

    # Поиск и перенос "УТ" в отдельный столбец
    for i, value in enumerate(row_data):
        if isinstance(value, str) and value.startswith("УТ"):
            ut_value = value
            row_data.pop(i)
            break

    # Поиск и перенос "Поступление" в отдельный столбец
    if isinstance(row_data[0], str) and row_data[0].startswith("Поступление"):
        current_receipt = row_data[0]  # Запоминаем поступление
        row_data.pop(0)

        # Извлекаем месяц и год из строки с Поступлением
        month_year_value = extract_month_year(current_receipt)

    if previous_ut_value:
        row_data.append(previous_ut_value)  # Перенос значения "УТ" на строку вверх
        previous_ut_value = None  # Сброс значения после переноса
    else:
        previous_ut_value = ut_value  # Запоминаем "УТ" для следующей строки
        row_data.append(None)  # Заполняем текущую строку пустым значением

    row_data.insert(0, current_receipt)  # Добавляем поступление в начало строки

    # Поиск кода из 7 цифр в формате X.XXX-XXX
    code_value = extract_code(row_data)

    return row_data, previous_ut_value, current_receipt, code_value, month_year_value

# Функция для извлечения кода из 7 цифр в формате X.XXX-XXX
def extract_code(row_data):
    for value in row_data:
        if isinstance(value, str):
            # Ищем код в формате X.XXX-XXX, где X - цифры
            match = re.search(r'\d{1,3}\.\d{3}-\d{3}', value)
            if match:
                return match.group(0)  # Возвращаем найденный код
    return None  # Если код не найден, возвращаем None

# Функция для извлечения месяца и года из строки "Поступление"
def extract_month_year(text):
    # Ищем дату в формате "ДД.ММ.ГГГГ"
    match = re.search(r'\d{2}\.\d{2}\.\d{4}', text)
    if match:
        date_str = match.group(0)
        # Извлекаем месяц и год
        day, month, year = date_str.split('.')
        return f"{month}.{year}"  # Возвращаем строку в формате ММ.ГГГГ
    return None  # Если дата не найдена, возвращаем None

# Функция для фильтрации данных
def filter_data(df):
    filtered_data = []
    codes = []  # Список для хранения найденных кодов
    months_years = []  # Список для хранения месяца и года
    previous_ut_value = None  # Переменная для хранения последнего значения "УТ"
    current_receipt = None  # Переменная для хранения последнего значения "Поступление"

    for row in df.itertuples(index=False):
        row_data, previous_ut_value, current_receipt, code_value, month_year_value = process_row(row, previous_ut_value, current_receipt)

        # Добавляем строку, если в первом столбце есть буквы
        if isinstance(row_data[1], str) and any(c.isalpha() for c in row_data[1]):
            filtered_data.append(row_data)
            codes.append(code_value)  # Добавляем код в список
            months_years.append(month_year_value)  # Добавляем месяц и год в список

    return filtered_data, codes, months_years

# Функция для удаления столбца, который содержит "УТ"
def remove_ut_column(df):
    # Ищем столбец, который содержит "УТ"
    for col in df.columns:
        if df[col].astype(str).str.contains("УТ").any():
            df = df.drop(columns=[col])  # Удаляем этот столбец
            break  # Прерываем, если нашли столбец с "УТ"
    return df

# Функция для сохранения данных в Excel
def save_to_excel(filtered_data, codes, months_years, output_file):
    filtered_df = pd.DataFrame(filtered_data)
    filtered_df['Code'] = codes  # Добавляем новый столбец с кодами
    filtered_df['Month_Year'] = months_years  # Добавляем новый столбец с месяцем и годом

    # Удаляем столбец с "УТ", если он существует
    filtered_df = remove_ut_column(filtered_df)

    # Не удаляем столбец "0", так как он нам нужен
    # filtered_df = filtered_df.drop(columns=[0])  # Удаляем только если нужно

    filtered_df.to_excel(output_file, index=False)
    print(f"Файл успешно сохранен как {output_file}")

# Основная функция
def main(file_path, output_file):
    df = load_data(file_path)  # Загрузка данных
    filtered_data, codes, months_years = filter_data(df)  # Фильтрация данных
    save_to_excel(filtered_data, codes, months_years, output_file)  # Сохранение в новый файл

# Запуск программы
if __name__ == "__main__":
    file_path = 'data_2.xlsx'  # Укажи свой файл
    output_file = 'converted_data_2.xlsx'  # Укажи имя файла для сохранения
    main(file_path, output_file)
