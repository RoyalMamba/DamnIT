import openpyxl

workbook= openpyxl.load_workbook('Files_and_Tags.xlsx')

sheet = workbook.active

filesnTags = [row for row in sheet.iter_rows(values_only = True , min_row = 2 )]

for fnt in filesnTags:
    print(fnt[1].split(','))