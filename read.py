import xlrd
book = xlrd.open_workbook('檔案名稱.xlsx') #開啟檔案
sheet = book.sheet_by_index(0) #根據順序獲取sheet頁
# sheet = book.sheet_by_name(0) #根據名稱獲取sheet頁
print(sheet.cell(0,2).value) #指定行和列獲取資料
print(sheet.cell(1,0).value) #指定行和列獲取資料
print(sheet.ncols) #獲取excel裡面有多少列
print(sheet.nrows) #獲取excel裡面有多少列
print(sheet.row_values(1))#取第幾欄的資料
print(sheet.col_values(0))#取第幾列的資料