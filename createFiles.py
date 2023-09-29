import os 
from openpyxl import Workbook
import re


def remove_special_characters(input_string):
    # Replace special characters with underscores
    clean_string = re.sub(r'[^a-zA-Z0-9]', '_', input_string)
    return clean_string

path = r'Assets\Anthony Albanese'

files = os.listdir(path)
files = sorted(files)
workbook = Workbook()
sheet = workbook.active

sheet['A1'] = 'FileNames'
sheet['B1'] = 'Tags'
for index, file in enumerate(files , start=2):
    try: 
        splitname, extension = os.path.splitext(file)
        cleaned_name = remove_special_characters(splitname)
        new_file_name = cleaned_name+extension
        oldPath = os.path.join(path,file)
        newPath = os.path.join(path,new_file_name)
        os.rename(oldPath,newPath)
        sheet[f'A{index}'] = new_file_name
    except:
        pass

excelFile = 'Files_and_Tags.xlsx'
workbook.save(excelFile)

