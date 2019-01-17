import openpyxl
import numpy
import xlrd
import xlwt

excel_path = r'../file/TaskDetail2059373187.xlsx'


def read_excel():
    global excel_path
    wb = xlrd.open_workbook(excel_path)
    sheet = wb.sheet_names()
    print(sheet)


class _Tasks:
    def __init__(self):
        self.Task = {'任务编号':'','需求提出方':'','需求编号':'','类型':'',}
        self.Tasks = []

    def Append_Rwbh(self,data):
        self.Task['任务编号'] = data[0]
        self.Task['需求提出方'] = data[1]
        self.Task['需求编号'] = data[2]
        self.Task['类型'] = data[3]
        self.Tasks.append(self.Task)


def main():
    data = ['1234','国泰君安','20183123','缺陷']
    task_data = _Tasks()
    task_data.Append_Rwbh(data)
    print(task_data.Tasks)


if __name__ == '__main__':
    main()