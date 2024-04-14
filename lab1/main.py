from constants.extensions import FileExtensions
from models.Book import Book
import os
import csv
import xml.etree.ElementTree as ET
from docx import Document
import pandas as pd
import tabula

# Общие глобальные переменные
delimiter = ';'
source_data_path = "./source_data"
result_data_path = "./result_data"

# Глобальный объект для хранения путей к файлам-источника
filepaths = {
    "txt": "",
    "xml": "",
    "docx": "",
    "xlsx": "",
    "pdf": ""
}

# Пути к результирующим файлам
results = {
    "txt": result_data_path + "/txt.csv",
    "xml": result_data_path + "/xml.csv",
    "docx": result_data_path + "/docx.csv",
    "xlsx": result_data_path + "/xlsx.csv",
    "pdf": result_data_path + "/pdf.csv"
}

# Проход по директории source_data для определения путей к файлам разных типов
try:
    # Движение по директории source_data
    for subdir, dirs, files in os.walk(source_data_path):
        # Перебор файлов
        for file in files:
            # Разбиение имени файла на отдельные части (расширение и всё остальное)
            elements = file.split('.')
            if len(elements) < 2:
                raise Exception("Error! The file does not contain the format")

            # Определение расширения файла
            fileExt = elements[len(elements) - 1]

            # Проверка расширения файла
            if fileExt in FileExtensions.as_list():
                filepaths[fileExt] = os.path.join(subdir, file)

except Exception as err:
    print(err)

# -----------------------------------------------------------------------------
# Convert TXT to CSV
# Открытие текстового файла
with open(filepaths["txt"], 'r') as in_file:
    # Получение всех строк из файла с преобразованием (удаление пробелов между частями таблиц)
    stripped = (' '.join(line.strip().split()) for line in in_file)

    # Формирование линий (строк таблицы)
    lines = (line.split(' ') for line in stripped if line)

    # Открытие результирующего файла и запись таблицы
    with open(results["txt"], 'w+') as out_file:
        writer = csv.writer(out_file, delimiter=delimiter)
        writer.writerows(lines)

# -----------------------------------------------------------------------------
# Convert XML to CSV
# Формирование первой строки таблицы, разбитой на категории хранимые в CSV-документе
dataForXml = [["Category", "Title", "Authors", "Year", "Price"]]

# Открытие XML-документа
with open(filepaths["xml"], 'r') as in_file:
    # Чтение всех данных из XML-документа
    xmlDocument = in_file.read()

    # Построение XML-дерева во внутреннем представлении на основе строки
    root = ET.ElementTree(ET.fromstring(xmlDocument))

    # Выбор всех дочерних элементов с тегом book (книга)
    for element in root.findall("book"):
        # Получение атрибута category из текущего тега book
        category = element.attrib.get('category')
        # Получение содержимого текста из тега title
        title = element.find("title").text
        # Получение содержимого текста из тега year
        year = int(element.find("year").text)
        # Получение содержимого текста из тега price
        price = float(element.find("price").text)

        authors = []

        # Получении информации об авторах книги
        for author in element.findall("author"):
            authors.append(author.text)

        # Формирование модели книги с помощью класса Book
        book = Book(category, title, authors, year, price)

        # Добавление данных в таблицу с конвертацией объекта book в строку
        dataForXml.append(book.to_list())

    # Открытие результирующего файла и запись таблицы
    with open(results["xml"], 'w+') as out_file:
        writer = csv.writer(out_file, delimiter=delimiter)
        writer.writerows(dataForXml)

# -----------------------------------------------------------------------------
# Convert DOCX to CSV
# Открытие документа DOCX
doc = Document(filepaths["docx"])

# Чтение всех таблиц из документа
all_tables = doc.tables

# Маркировка таблиц в структуре
data_tables = {i: None for i in range(len(all_tables))}

# Прохождение в цикле по всем таблицам
for i, table in enumerate(all_tables):
    # Добавление подмассивов для значений ячеек таблиц
    data_tables[i] = [[] for _ in range(len(table.rows))]

    # Добавление значений из ячеек прочитанной таблицы в DOCX
    for j, row in enumerate(table.rows):
        for cell in row.cells:
            data_tables[i][j].append(cell.text)

# Открытие результирующего файла и запись таблицы
with open(results["docx"], 'w+') as out_file:
    writer = csv.writer(out_file)

    for table in data_tables[0]:
        writer.writerow(table)

# -----------------------------------------------------------------------------
# Convert XLSX to CSV
# Чтение EXCEL-документа
df = pd.read_excel(filepaths["xlsx"], sheet_name="PlayerData")

# Удаляем индексный столбец
df.drop(index=df.index[0], axis=0, inplace=True)

# Запись результата в новый документ (результат)
df.to_csv(results["xlsx"], index=False, header=False)

# -----------------------------------------------------------------------------
# Convert PDF to CSV

# Добавление переменной окружения для поддержки jpype
os.environ["JAVA_HOME"] = "C:\\Program Files\\Microsoft\\jdk-11.0.16.101-hotspot"

# Чтение pdf таблицы
pages = tabula.read_pdf(filepaths["pdf"], pages="all")

# Ручное задание заголовков таблицы в данных о таблице (9 столбцов)
data_table = [
    [
     "Row Labels",
     "Average of Date",
     "Min of Hour",
     "Max of Tx",
     "Min of Tn",
     "Average off ff-mean",
     "Max of ff-gust",
     "Sum of Rainfall",
     "Max of RainRate"
     ]
]


# Проход по всем определённым страницам документа, в которых есть таблица
for index in range(len(pages)):
    # Получение информации об одной таблице
    page = pages[index]

    # Удаление из столбцов некорректных данных
    page.columns = page.columns.str.replace('-.1', '-')

    # Создание копии-объекта без столбцов с именем Unnamed (нет имени)
    page_without_unnamed = page.loc[:, ~page.columns.str.contains('^Unnamed')]

    # Если index == 0, то удаляем столбец со значением NaN (неопределённым)
    if index == 0:
        page = page.dropna(axis=1)
    else:
        # Иначе заполняем все значения NaN пустой строкой
        page = page.fillna('')

    # Добавление строк, характеризующие подтаблицы в рамках одной страницы
    data_table.append(list(page_without_unnamed.columns))

    # Последовательное добавление элементов распознанных данных в таблицу
    for item in page.values:
        data_table.append(item)

# Открытие результирующего файла и запись таблицы
with open(results["pdf"], 'w+') as out_file:
    writer = csv.writer(out_file, delimiter=delimiter)

    for table in data_table:
        writer.writerow(table)