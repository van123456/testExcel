#!/usr/bin/env python
# -*-coding:utf-8-*-
# exceltest.py
import xlrd


def open_excel(file=r"C:\Users\van\Desktop\冒烟测试用例.xlsx"):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as err:
        print(err)


def openexcel_sheetbyname(sheet_name="Sheet1", index=1):
    data = open_excel()
    # 打开具体的sheet表
    sheet = data.sheet_by_name(sheet_name)
    # 我想要获取行数
    nrows = sheet.nrows
    # 获取具体的行的值
    value1 = sheet.row_values(index)
    list = []
    # 遍历所有行
    for row in range(nrows):
        value = sheet.row_values(row)
        # print(value)
        if value:
            list1 = {}
            # 遍历第一行的个数
            for n in range(len(value1)):
                # (list[row]:value)
                list1[value1[n]] = value[n]
            list.append(list1)
    return list


def operation():
    table = openexcel_sheetbyname()
    for n in table:
        # print(n)
        for value in n.values():
            print(value)


if __name__ == '__main__':
    operation()
