import xlrd
import xlsxwriter
import pandas as pd

df = pd.read_excel(r"C:\Users\M06429\Desktop\Portable Python-3.8.2 x64\test\test1.xlsx", sheet_name="工作表1", index_col=None, header = None,skiprows = 1)
df

df2 = pd.read_excel(r"C:\Users\M06429\Desktop\Portable Python-3.8.2 x64\test\test2.xlsx", sheet_name="工作表1", index_col=None, header=None, skiprows = 1)
df2

dfs = [df,df2]

df3 = pd.concat(dfs, sort=False)
df3.to_excel(r"C:\Users\M06429\Desktop\Portable Python-3.8.2 x64\test\test5.xlsx",header = False,index = False)
