import xlrd
import xlsxwriter
import xlrd
from datetime import date, datetime


import pyodbc
cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\房貸商機清單.accdb')
crsr = cnxn.cursor()

workbook = xlrd.open_workbook(r"\\10.240.172.69\r52500_電話行銷科tma行政共用\MIS REPORT\案件回報\房貸商機清單.xlsx")

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

print(data_list)
#print("INSERT INTO 尚未進二段清單_回報 VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"%(data_list[i][0],data_list[i][1],data_list[i][2],data_list[i][3],data_list[i][4],data_list[i][5],data_list[i][6],data_list[i][7],data_list[i][8],data_list[i][9]))



#for i in range(len(data_list)):
    #crsr.execute("INSERT INTO 尚未進二段清單_回報 VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"%(data_list[i][0],data_list[i][1],data_list[i][2],data_list[i][3],data_list[i][4],data_list[i][5],data_list[i][6],data_list[i][7],data_list[i][8],data_list[i][9]))

#crsr.commit()
#crsr.close()
#cnxn.close()
