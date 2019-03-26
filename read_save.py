# coding = utf-8
import xlrd
import xlwt
import os

file = "D:\Pycharm\PyCharm 2018.2.4\Recomend-system\SELECT\食谱\\"
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('recipes')
j = 0
k = 0
for classname in os.listdir(file):
    xls = os.path.join(file, classname)
    for xlsname in os.listdir(xls):
        xlspath = os.path.join(xls, xlsname)
        print(xlspath)
        data = xlrd.open_workbook(xlspath)
        # sheetname提取
        sheets = data.sheet_names()
        for sheet_name in sheets:
            # 得到每个sheet对应的excel表
            table = data.sheet_by_name(sheet_name)
            nrows = table.nrows
            ncols = table.ncols
            # 得到表中第一列的值
            # recipes = table.col(0)
            for i in range(nrows):
                recipename = table.col(0)[i].value
                # if recipename not in newrecipes:
                # 向新的Excel中写入数据
                for j in range(ncols):
                    worksheet.write(k, j, table.row_values(i)[j])
                k += 1
                workbook.save('食谱.xls')

