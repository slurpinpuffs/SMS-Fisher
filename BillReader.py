import PyPDF2
from openpyxl import Workbook
import re

file = input("Please enter the filepath for the bill you'd like to input (no quotation marks): ")
folder = input("Please enter the filepath for the folder you'd like the outputted sheet to go to (no quotation marks): ")
page_list = []
reader = PyPDF2.PdfReader(file)
count = len(reader.pages)

workbook = Workbook()
sheet = workbook.active
row = 1

for page_number in range(count):
    page = reader.pages[page_number]
    page_list.append(page)

for page in page_list:
    page_content = page.extract_text()
    print(page_content)
    for name in re.findall(r'\b9C\S*', page_content):
        sheet[f'A{row}'] = name
        sheet[f'B{row}'] = '1'
        print(f'A{row}')
        row += 1

workbook.save(filename=folder + r'\bill.xlsx')


