# coding = utf-8
import pandas as pd


# 获取指定的表单
df = pd.read_excel('食谱.xlsx') #默认为第一个sheet
table = pd.read_excel('食谱.xlsx', sheet_name='recipes')
nrows=df.shape[0]
ncols=df.columns.size
# 获取某一行数据 data = df.ix[0].values(很少使用了)
data = df.loc[0].values
# 获取多行多列的数据
datas = df.ix[:, 0]
# 去除重复行
df.drop_duplicates(subset=['A'], keep='first')
# 去除指定（重复）行
recipes = []
k = 0
for i in range(nrows):
    # 根据位置查找元素
    recipe = df.iat[i, 0]
    if recipe not in recipes:
        recipes.append(recipe)
        k += 1
        print(recipe)
    else:
        df.drop(i)
print(k)


