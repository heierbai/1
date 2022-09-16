import xlrd
import xlwt
import datetime

""" 打开excel表格 """
workbook = xlrd.open_workbook("1.xlsx")
print(workbook)       # 结果：<xlrd.book.Book object at 0x000000000291B128>
 
 
""" 获取所有sheet名称 """
sheet_names = workbook.sheet_names()
print(sheet_names)     # 结果：['表1', 'Sheet2']



 
# 通过name获取第一个sheet对象
sheet1_object = workbook.sheet_by_name(sheet_name="Sheet1")
print(sheet1_object)    # 结果：<xlrd.sheet.Sheet object at 0x0000000002956710>
# 通过name获取第二个sheet对象
sheet2_object = workbook.sheet_by_name(sheet_name="Sheet2")
print(sheet2_object)    # 结果：<xlrd.sheet.Sheet object at 0x0000000002956710>
 

# 通过sheet名称判断sheet1是否导入
sheet1_is_load = workbook.sheet_loaded(sheet_name_or_index="Sheet1")
print(sheet1_is_load)    # 结果：True
sheet2_is_load = workbook.sheet_loaded(sheet_name_or_index="Sheet2")
print(sheet2_is_load)    # 结果：True


""" 对sheet对象中的行执行操作 """
# 获取sheet1中的有效行数
nrows1 = sheet1_object.nrows
print(nrows1)        # 结果：
# 获取sheet2中的有效行数
nrows2 = sheet2_object.nrows
print(nrows2)        # 结果：

""" 对sheet对象中的列执行操作 """
# 获取sheet1中的有效列数
ncols1 = sheet1_object.ncols
print(ncols1)        # 结果：
# 获取sheet2中的有效列数
ncols2 = sheet2_object.ncols
print(ncols2)        # 结果：


# 获取sheet1中第rowx=1行，第colx=2列的单元值
#cell_value = sheet1_object.cell_value(rowx=0, colx=1)
#print(cell_value)      # 结果: 



# 创建一个workbook 设置编码
workbook2 = xlwt.Workbook(encoding='utf-8')
 
 
# 创建一个worksheet
worksheet2 = workbook2.add_sheet('Sheet1')

#字体样式设置
style = xlwt.XFStyle()  # 初始化样式
font = xlwt.Font()  # 为样式创建字体
font.name = 'Times New Roman'
font.height = 20 * 11  # 字体大小，11为字号，20为衡量单位
font.bold = False  # 黑体
font.underline = False # 下划线
font.italic = False # 斜体字
style.font = font # 设定样式

# 数据写入excel,参数对应 行, 列, 值
#worksheet.write(0, 0, 'test_data') # 不带样式的写入


# 获取sheet1中第colx=0列的数据
col_values1 = sheet1_object.col_values(colx=0)
col_values2 = sheet2_object.col_values(colx=0)
i=0
for x in col_values1:
    for y in col_values2:
        if x == y :
            print(x)
            worksheet2.write(i, 0, x, style) # 带字体样式的写入
            i=i+1



# 保存文件
workbook2.save('2.xls')