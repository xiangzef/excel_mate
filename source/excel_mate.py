import openpyxl
import numpy
import xlrd
import xlwt



excel_path = r'../file/TaskDetail145422954.xlsx'

wb = xlrd.open_workbook(excel_path)
sheet = wb.sheet_names()
print(sheet)

# def genDictByPhone(xls_file,phone):
#     l = []
#     wb = xlrd.open_workbook(xls_file)
#     sheet = wb.sheet_by_index(0)
#     for irow in range(sheet.nrows):
#         c_row = sheet.row(irow)
#         if str(c_row[3].value) == str(phone):
#             d = {}
#             d['name'] = c_row[0].value
#             d['age'] = c_row[1].value
#             d['addr'] = c_row[2].value
#             d['phone'] = c_row[3].value
#             l.append(d)
#     return l
#
# genDictByPhone(excel_path, '缺陷')