import xlwings as xw

wb = xw.Book("副本share_file.xlsx")
sht = wb.sheets['工作表1']
pos = ""
content = []
for i in range(1,3091):
    pos = "A" + str(i)
    content.append(sht.range(pos).value)

for n in content:
    print(n)
    
wb.close()
