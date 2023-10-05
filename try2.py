#import 
import xlwings as xw


#value
wb = xw.Book()
sht = wb.sheets['sheet1']
A_pos = 'A'
pos = 0
#loop
for i in range(1,101):
    pos = A_pos + str(i)
    sht.range(pos).value = i
    print(i)

wb.save('Data_Try.xlsx')
wb.close()
