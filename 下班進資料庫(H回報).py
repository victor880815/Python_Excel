import openpyxl
import xlrd
import xlsxwriter
import xlwt
import pandas as pd
from datetime import date, datetime

#合併H回報
df = pd.read_excel(r"\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\未送二段原因回報_邱秀環.xlsx", sheet_name="H回報", index_col=None, header = 0,skiprows = 0)
#print(df)


df2 = pd.read_excel(r"\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\未送二段原因回報_郭曉清.xlsx", sheet_name="H回報", index_col=None, header= 0, skiprows = 0)
#print(df2)

dfs = [df,df2]

df3 = pd.concat(dfs, sort=False)
df3.to_excel(r"\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\書翊\H回報.xlsx",sheet_name="工作表1", index = False ,header = 0)

#print(df3)



a = df3.drop_duplicates(subset=["申請書編號"], keep="last")
#print(a)

from pandas import DataFrame
df4 = DataFrame(a)
df4.to_excel("Z:\MIS REPORT\案件回報\書翊\H回報.xlsx",sheet_name="工作表1", index = False ,header = 0)
#print(df4)

#H回報丟入資料庫
import pyodbc
cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\信貸案件進度回報.accdb')
crsr = cnxn.cursor()

workbook = xlrd.open_workbook("Z:\MIS REPORT\案件回報\書翊\H回報.xlsx")

sheet = workbook.sheet_by_index(0)
data_list = []
row_list = []
nRows = sheet.nrows
nCols = sheet.ncols
for i in range(0, nRows):
    row_list = []
    for j in range(nCols):
        # 獲取第i行，第j列的值
        data_value = sheet.cell(i, j).value
        # 獲取第i行，第j列的型別
        # ctype :  0 empty, 1 string ,2 number, 3 date, 4 boolean 5,error
        data_type = sheet.cell(i, j).ctype
        if data_type == 2:
            # 將字串轉為number
            data_value = str(int(data_value))
        if data_type == 3:
            # 對讀取資料表中日期列 進行格式化
            date_t = xlrd.xldate_as_tuple(data_value, workbook.datemode)
            data_value = date(*date_t[:3]).strftime('%Y-%m-%d')
        row_list.append(data_value)
    data_list.append(row_list)

#print(data_list)
#print("INSERT INTO 尚未進二段清單_回報 VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"%(data_list[i][0],data_list[i][1],data_list[i][2],data_list[i][3],data_list[i][4],data_list[i][5],data_list[i][6],data_list[i][7],data_list[i][8],data_list[i][9]))



for i in range(len(data_list)):
    crsr.execute("INSERT INTO 尚未進二段清單_回報 VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"%(data_list[i][0],data_list[i][1],data_list[i][2],data_list[i][3],data_list[i][4],data_list[i][5],data_list[i][6],data_list[i][7],data_list[i][8],data_list[i][9]))

crsr.commit()
crsr.close()
cnxn.close()


