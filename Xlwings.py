import xlwings as xw

wb = xw.Book()
sht = wb.sheets['sheet1']
sht.range('A1').value = '123'
sht.range('A1').value
wb.save('First.xlsx')
wb.close()
