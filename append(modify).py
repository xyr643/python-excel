import xlrd
from xlutils.copy import copy
import random


# 随机选取Excel数据并打上-标签
def generate_neglabel(table, xlwrite, sheet, nrows, ncols, label):
    # 存储随机数（无重复）
    index = set()
    while len(index) < 100:
        temp = random.randint(0, nrows)
        index.add(temp)
        lb = table.col_values(ncols-2)[temp]
        print(temp, lb)
        if not lb:
            sheet.write(temp, ncols-2, label)
            sheet.write(temp, ncols-1, '-')
        else:
            labels = lb + ' ' + lable
            sheet.write(temp, ncols-2, labels)
    xlwrite.save('标签.xls')

# 随机选取Excel数据并打上+标签
def generate_poslabel(table, xlwrite, sheet, ncols, label, rec):
    # 存储随机数（无重复）
    list = random.sample(rec, 300)
    for temp in list:
        print(temp)
        lb = table.col_values(ncols - 2)[temp]
        if not lb:
            sheet.write(temp, ncols-2, label)
            sheet.write(temp, ncols-1, '+')
        else:
            labels = lb + ' ' + label
            sheet.write(temp, ncols-2, labels)
    xlwrite.save('标签.xls')


# 将Excel多列合并到一列（运行速度较慢，建议pandas）
def concate_cols(table, xlwrite, sheet, nrows, ncols):
    for i in range(nrows):
        tag = ''
        for j in range(32, 39):
            value = table.col(j)[i].value
            if not value:
                continue
            tag = tag + value + ' '
        print(i)
        sheet.write(i, ncols, tag)
        xlwrite.save('标签.xls')

def Main():
    # 读取原始excel
    data = xlrd.open_workbook('D:\Pycharm\PyCharm 2018.2.4\Recomend-system\SELECT\标签.xls')
    table = data.sheet_by_name('recipes')
    nrows = table.nrows
    ncols = table.ncols
    print(nrows, ncols)

    # 修改原始Excel
    xlwrite = copy(data)
    # 定位到目标sheet
    sheet = xlwrite.get_sheet(0)
    '''
    进行以下操作需先对Excel中存储标签的两列进行初始化
    这里是单独操作label_0,再for循环操作以下label
    '''
    neglabels = ['label1', 'label2', 'label3']
    poslabels = ['label4', 'label5', 'label6']
    for lable in neglabels:
        generate_neglabel(table, xlwrite, sheet, nrows, ncols, label)

    rec = []
    for i in range(nrows):
        lable = table.col_values(ncols-1)[i]
        if lable == '-':
            continue
        else:
            rec.append(i)
    for lable in poslabels:
        generate_poslabel(table, xlwrite, sheet, ncols, label, rec)
    #concate_cols(table, xlwrite, sheet, nrows, ncols)

if __name__ == '__main__':
    Main()



