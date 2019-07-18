import docx , os
import openpyxl
from docx import Document

file_name =  os.getcwd()+'\\' + '附件12.2019年省级工业互联网发展专项资金申报指南.docx'
doc = docx.Document(file_name)
paragraphs = doc.paragraphs
print(len(paragraphs))
print(paragraphs[4].text)
document = Document(file_name)
tables = document.tables
print(len(tables))
table = tables[0]
for i in range(0,len(table.rows)):
    row_cells = table.rows[i].cells
    for cell in row_cells:
        print(cell.text)

