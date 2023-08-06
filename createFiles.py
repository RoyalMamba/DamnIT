import os 
from openpyxl import Workbook


files = os.listdir(r'Assets\Fumio Kishida')
files = sorted(files)
workbook = Workbook()
sheet = workbook.active

sheet['A1'] = 'FileNames'
sheet['B1'] = 'Tags'

for index, file in enumerate(files , start=2):
    sheet[f'A{index}'] = file

excelFile = 'Files_and_Tags.xlsx'
workbook.save(excelFile)

