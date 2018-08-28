#!/usr/bin/env python
# coding='utf-8'
# 写入excel文件
import xlwt
import os

raw_Projectname = input('请输入你的项目名称：')
wbk = xlwt.Workbook()
# 新建sheet表
sheet = wbk.add_sheet('测试用例', cell_overwrite_ok=True)

# 设置style颜色

# 设置调色版
# 红色
xlwt.add_palette_colour('style0', 60)
# 绿色
xlwt.add_palette_colour('style1', 50)
# 黑色
xlwt.add_palette_colour('style2', 8)
# 设置style1-调色、加粗
# 图案,
style0 = xlwt.easyxf('font:colour style0 ')
style1 = xlwt.easyxf('font: colour style2, bold True;'
                     'pattern: pattern solid, fore_colour style1,back_colour black;'
                     'alignment: horz center,vert center'
                     )

# 设置行宽
sheet.col(1).width = 256 * 20

# 合并单元格
# style = xlwt.XFStyle()
# 第2行-第2行,第3列-第10列
sheet.write_merge(1, 1, 2, 9, '%s' % raw_Projectname)
# sheet.write_merge(1, 1, 2, 9, '测试项目')
sheet.write_merge(2, 2, 2, 9, '测试的功能模块')
sheet.write_merge(3, 3, 2, 9, 'van')
sheet.write_merge(4, 4, 2, 9, 'van')
sheet.write_merge(5, 5, 2, 9, '测试模块：具体包括1、2、3')

# 单元格的写入
# 写入第二行,第二列,
sheet.write(1, 1, '项目名称', style1)
sheet.write(2, 1, '功能模块', style1)
sheet.write(3, 1, '用例编写人', style1)
sheet.write(4, 1, '用例执行人', style1)
sheet.write(5, 1, '用例说明', style1)
sheet.write(6, 1, '测试模块', style1)
sheet.write(6, 2, '测试点', style1)
sheet.write(6, 3, '测试分支', style1)
sheet.write(6, 4, '用例描述', style1)
sheet.write(6, 5, '预期结果', style1)
sheet.write(6, 6, '实际结果', style1)
sheet.write(6, 7, 'BUG', style1)
sheet.write(6, 8, '备注', style1)
sheet.write(6, 9, '归属', style1)

'''
# 测试用例（直接录入）
raw_Module = input("请输入你的测试模块:")
sheet.write_merge(7, 7, 1, 1, '%s' % raw_Module)
raw_Point = input("请输入你的测试点:")
sheet.write_merge(7, 7, 2, 2, '%s' % raw_Point)
raw_Branch = input("请输入你的测试分支:")
sheet.write_merge(7, 7, 3, 3, '%s' % raw_Branch)
'''
# 测试用例，交互式录入
# 控制开关
running = True
n = 7
while running:
    raw_Module = input("请输入你的测试模块")
    sheet.write_merge(n, n, 1, 1, '%s' % raw_Module)
    raw_Point = input("请输入你的测试点")
    sheet.write_merge(n, n, 2, 2, '%s' % raw_Point)
    raw_Branch = input("请输入你的测试分支")
    sheet.write_merge(n, n, 3, 3, '%s' % raw_Branch)
    raw_Script = input("请输入你的用例描述")
    sheet.write_merge(n, n, 4, 4, '%s' % raw_Script)
    raw_Expectedoutcome = input("请输入你的预期结果")
    sheet.write_merge(n, n, 5, 5, '%s' % raw_Expectedoutcome)
    raw_Actualoutcome = input("请输入你的实际结果")
    sheet.write_merge(n, n, 6, 6, '%s' % raw_Actualoutcome)
    raw_Bug = input("请输入BUG")
    sheet.write_merge(n, n, 7, 7, '%s' % raw_Bug)
    raw_Remarks = input("请输入备注")
    sheet.write_merge(n, n, 8, 8, '%s' % raw_Remarks)
    raw_Attribution = input("请输入归属")
    sheet.write_merge(n, n, 9, 9, '%s' % raw_Attribution)
    n = n + 1
    raw_Final = input("录入结束了吗（Y or N）")
    if raw_Final == 'Y' and 'y':
        print("测试用例录入结束了")
        running = False
    else:
        print("继续录入中")

os.makedirs(r'C:/Users/xiewenhua/Desktop/Testexcel')
# xlsx表的保存
wbk.save(r'C:\Users\xiewenhua\Desktop\Testexcel\test.xls')
