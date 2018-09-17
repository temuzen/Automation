
from xlwings import Book, Sheet, Range, Chart
import xlwings as xw
import time

book1 = xw.Book()
sht = xw.books('book1').sheets[1]
wrt = xw.books('book1').sheets[2]

xx=[] 
'''
def getQuote(ticker, type):
    Range('B1').value = '=RTD("pi.rtdserver", ,"%s", "%s")' % (type, ticker)
    time.sleep(2.5)
    x = Range('B1').value
    return Range('B1').value

getQuote("Volume","MCX_ALUMINI18SEPFUT")

'''

sht.range('B1').value = '=RTD("pi.rtdserver",,"MCX_ALUMINI18SEPFUT","TradingSymbol")'
sht.range('C1').value = '=RTD("pi.rtdserver",,"MCX_ALUMINI18SEPFUT","Last")'
sht.range('D1').value = '=RTD("pi.rtdserver",,"MCX_ALUMINI18SEPFUT","BidSize")'
sht.range('E1').value = '=RTD("pi.rtdserver",,"MCX_ALUMINI18SEPFUT","AskSize")'
sht.range('F1').value = '=RTD("pi.rtdserver",,"MCX_ALUMINI18SEPFUT","Bid")'
sht.range('G1').value = '=RTD("pi.rtdserver",,"MCX_ALUMINI18SEPFUT","Ask")'
sht.range('H1').value = '=RTD("pi.rtdserver",,"MCX_ALUMINI18SEPFUT","lastUpdateTime")'
   
   
import time
for i in range(0,1000):

    x0 = str(sht.range('B1').value)
    x1 = str(sht.range('C1').value)
    x2 = str(sht.range('D1').value)
    x3 = str(sht.range('E1').value)
    x4 = str(sht.range('F1').value)
    x5 = str(sht.range('G1').value)
    x6 = sht.range('H1').value
    
    
    wrt.range('A'+str(i+1)).value = x0   
    wrt.range('B'+str(i+1)).value = x1   
    wrt.range('C'+str(i+1)).value = x2   
    wrt.range('D'+str(i+1)).value = x3   
    wrt.range('E'+str(i+1)).value = x4   
    wrt.range('F'+str(i+1)).value = x5 
    wrt.range('G'+str(i+1)).value = x6 

     
    print(x0,x1,x2,x3,x4,x5,x6)
  
    sleep(1)












