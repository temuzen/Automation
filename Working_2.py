
from xlwings import Book, Sheet, Range, Chart
import xlwings as xw
import time


wb = xw.Book(r'C:\Users\Yash\Desktop\Marketwatch.xlsx')

x=['']*20

for i in range(2,20):
    x[i] = str(wb.sheets['sheet'].range('B'+str(i)).value)
    print(x[i])
    Sheetname[i] = x[i]
    wb.sheets(Sheetname[i])
    
  

write1 = xw.books('book1').sheets[2]
write2 = xw.books('book1').sheets[3]
write3 = xw.books('book1').sheets[4]
write4 = xw.books('book1').sheets[5]
write5 = xw.books('book1').sheets[6]








   




counter = [1,1,1,1,1]

a = counter[0]
b = counter[1]
c = counter[2]
d = counter[3]
e = counter[4]



while(1):    


def Writer(Sheetno,counterno):

    x0 = str(sheetno.range('B2').value)
    x1 = str(sheetno.range('C2').value)
    x2 = str(sheetno.range('D2').value)
    x3 = str(sheetno.range('E2').value)
    x4 = str(sheetno.range('F2').value)
    x5 = str(sheetno.range('G2').value)
    x6 = str(sheetno.range('H2').value)
       
    String1 = x0+x1+x2+x3+x4+x5+x6 
     
      print(String1)
    
    y0 = str(sheetno.range('B2').value)
    y1 = str(sheetno.range('C2').value)
    y2 = str(sheetno.range('D2').value)
    y3 = str(sheetno.range('E2').value)
    y4 = str(sheetno.range('F2').value)
    y5 = str(sheetno.range('G2').value)
    y6 = str(sheetno.range('H2').value)

    String2 = y0+y1+y2+y3+y4+y5+y6 

    if String1==String2:
        break
    else:
        wrt.range('A'+str(i+1)).value = x0   
        wrt.range('B'+str(i+1)).value = x1   
        wrt.range('C'+str(i+1)).value = x2   
        wrt.range('D'+str(i+1)).value = x3   
        wrt.range('E'+str(i+1)).value = x4   
        wrt.range('F'+str(i+1)).value = x5 
        wrt.range('G'+str(i+1)).value = x6 

        counterno=counterno+1





'''