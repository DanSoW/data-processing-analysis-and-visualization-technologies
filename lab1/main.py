from constants.extensions import FileExtensions
from models.Book import Book
import os
import csv
import xml.etree.ElementTree as ET
from docx import Document
import pandas as pd

delimiter = ';'
source_data_path = "./source_data"
result_data_path = "./result_data"

filepaths = {
    "txt": "",
    "xml": "",
    "docx": "",
    "xlsx": "",
    "pdf": ""
}

results = {
    "txt": result_data_path + "/txt.csv",
    "xml": result_data_path + "/xml.csv",
    "docx": result_data_path + "/docx.csv",
    "xlsx": result_data_path + "/xlsx.csv"
}

# Get all filepaths
try:
    for subdir, dirs, files in os.walk(source_data_path):
        for file in files:
            elements = file.split('.')
            if len(elements) < 2:
                raise Exception("Error! The file does not contain the format")

            fileExt = elements[len(elements) - 1]

            if fileExt in FileExtensions.as_list():
                filepaths[fileExt] = os.path.join(subdir, file)

except Exception as err:
    print(err)

# -----------------------------------------------------------------------------
# Convert TXT to CSV
with open(filepaths["txt"], 'r') as in_file:
    stripped = (' '.join(line.strip().split()) for line in in_file)
    lines = (line.split(' ') for line in stripped if line)

    with open(results["txt"], 'w+') as out_file:
        writer = csv.writer(out_file, delimiter=delimiter)
        writer.writerows(lines)

# -----------------------------------------------------------------------------
# Convert XML to CSV
dataForXml = [["Category", "Title", "Authors", "Year", "Price"]]

with open(filepaths["xml"], 'r') as in_file:
    xmlDocument = in_file.read()
    root = ET.ElementTree(ET.fromstring(xmlDocument))

    for element in root.findall("book"):
        category = element.attrib.get('category')
        title = element.find("title").text
        year = int(element.find("year").text)
        price = float(element.find("price").text)

        authors = []
        for author in element.findall("author"):
            authors.append(author.text)

        book = Book(category, title, authors, year, price)
        dataForXml.append(book.to_list())

    with open(results["xml"], 'w+') as out_file:
        writer = csv.writer(out_file, delimiter=delimiter)
        writer.writerows(dataForXml)

# -----------------------------------------------------------------------------
# Convert DOCX to CSV
doc = Document(filepaths["docx"])

all_tables = doc.tables
data_tables = {i:None for i in range(len(all_tables))}

for i, table in enumerate(all_tables):
    data_tables[i] = [[] for _ in range(len(table.rows))]

    for j, row in enumerate(table.rows):
        for cell in row.cells:
            data_tables[i][j].append(cell.text)

with open(results["docx"], 'w+') as out_file:
    writer = csv.writer(out_file)

    for table in data_tables[0]:
        writer.writerow(table)

# -----------------------------------------------------------------------------
# Convert XLSX to CSV
df = pd.read_excel(filepaths["xlsx"], sheet_name="PlayerData")
df.drop(index=df.index[0], axis=0, inplace=True)
print(df)
df.to_csv(results["xlsx"], index=False, header=False)
