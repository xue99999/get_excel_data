from openpyxl import load_workbook
from openpyxl import Workbook

from datetime import datetime

def get_sheet_data(wb, sheet_item):
    sheet = wb.get_sheet_by_name(sheet_item)
    sheet_row = sheet.max_row
    sheet_column = sheet.max_column
    data_lst = []

    for row in range (2, sheet_row+1):
        obj = {}
        date = sheet['A' + str(row)].value
        name = sheet['B' + str(row)].value
        num  = sheet['C' + str(row)].value
        obj = { 'date': date, 'name': name, 'num': num}
        data_lst.append(obj)

    return data_lst


def combine():
    wb = load_workbook('courses.xlsx')
    lst = wb.get_sheet_names()
    sheet1_data = get_sheet_data(wb, lst[0])
    sheet2_data = get_sheet_data(wb, lst[1])
    print(sheet1_data[:3])
    print(sheet2_data[:3])

def split():
    pass 

if __name__ == '__main__':
    combine() 
    split()
