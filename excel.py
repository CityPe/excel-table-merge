# -*- coding: utf-8 -*-

import xlrd
import os
import xlsxwriter

PROJECT_PATH = "/Users/xihe/Desktop"
SAVE_PATH = "/Users/xihe/Desktop/excel.xlsx"

# 扫描文件夹
def scan_files(rootdir):
    files = []
    for parent, dirnames, filenames in os.walk(rootdir):
        for filename in filenames:
            files.append(os.path.join(parent,filename))
    return files
    
# 打开文件
def open_xls(path):
    file = xlrd.open_workbook(path)
    return file

# 获取所有 sheet 如果不是只存在一张 sheet 
def get_sheet(file):
    return file.sheets()

# 获取 sheet 表行数
def get_nrows(file,sheet_index):
    table = file.sheets()[sheet_index]
    return table.nrows

# 读取文件内容并返回
def getfile_content(path,sheet_index):
    file = open_xls(path)
    table = file.sheets()[sheet_index]
    total_rows = table.nrows
    for row in range(total_rows):
        row_data = table.row_values(row)
        datas.append(row_data)
    return datas

# 读取 sheet 表的个数
def getsheet_num(file):
    count = 0
    sheets = get_sheet(file)
    for sheet in sheets:
        count += 1
    return count

# 生成新文件
def generate_file(rvalue):
    endfile = SAVE_PATH
    workbook = xlsxwriter.Workbook(endfile)

    # 创建一个 sheet 工作对象
    worksheet = workbook.add_worksheet()
    for i in range(len(rvalue)):
        for x in range(len(rvalue[i])):
            item = rvalue[i][x]
            worksheet.write(i,x,item)
    workbook.close()
    print(" successed!! ".center(40,"🍺"))

if __name__ == '__main__':
    # 扫描要合并的 excel 文件列表
    files = scan_files(PROJECT_PATH)
    allxle_paths = list(filter(lambda filename: filename.endswith('xlsx'),files))
    
    # 存储所有读取的结果
    datas = []
    for path in allxle_paths:
        file = open_xls(path)
        count = getsheet_num(file)
        for sheet_index in range(count):
            print("正在读取文件:" +str(path) + "的第" + str(sheet_index) + "个 sheet 表....")
            rvalue = getfile_content(path,sheet_index)
    
    # 生成最终文件
    generate_file(rvalue)
    