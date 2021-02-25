import openpyxl
import xlrd
import xlsxwriter
import xlwt
import pandas as pd
df = pd.read_excel(r"\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\未送二段原因回報_邱秀環.xlsx", sheet_name="H回報", index_col=None, header = None)
#print(df)

a = df.drop_duplicates(subset=['戶名'], keep='last')
print(a)


from pandas import DataFrame
df = DataFrame(a)
df.to_excel(r"C:\Users\M06429\Desktop\123.xlsx", header = False , index = False)
print(df)