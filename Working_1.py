
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
'''

def Quote(ticker1):    
    xlwings.books(str(ticker1)).sheets[1]

    sheet.range('A1').value = '"TradingSymbol")'
    sheet.range('B2').value = '=RTD("pi.rtdserver",,"%s","TradingSymbol")'
    sheet.range('C2').value = '=RTD("pi.rtdserver",,"%s","Last")'
    sheet.range('D2').value = '=RTD("pi.rtdserver",,"%s","BidSize")'
    sheet.range('E2').value = '=RTD("pi.rtdserver",,"%s","AskSize")'
    sheet.range('F2').value = '=RTD("pi.rtdserver",,"%s","Bid")'
    sheet.range('G2').value = '=RTD("pi.rtdserver",,"%s","Ask")'
    sheet.range('H2').value = '=RTD("pi.rtdserver",,"%s","lastUpdateTime")'
    
    
def write_to(ticker,sheet):    
    xlwings.books(str(ticker)).sheets[1]
    sheet.range('B2').value = '=RTD("pi.rtdserver",,"%s","TradingSymbol")'
    sheet.range('C2').value = '=RTD("pi.rtdserver",,"%s","Last")'
    sheet.range('D2').value = '=RTD("pi.rtdserver",,"%s","BidSize")'
    sheet.range('E2').value = '=RTD("pi.rtdserver",,"%s","AskSize")'
    sheet.range('F2').value = '=RTD("pi.rtdserver",,"%s","Bid")'
    sheet.range('G2').value = '=RTD("pi.rtdserver",,"%s","Ask")'
    sheet.range('H2').value = '=RTD("pi.rtdserver",,"%s","lastUpdateTime")'
    
''' 

symbol='"NFO_BANKNIFTY18SEPFUT"'
#                        =RTD("pi.rtdserver", ,"NFO_BANKNIFTY18SEPFUT", "TradingSymbol")
sht.range('B2').value = '=RTD("pi.rtdserver", ,'+symbol+', "TradingSymbol")'
sht.range('C2').value = '=RTD("pi.rtdserver", ,'+symbol+', "Last")'
sht.range('D2').value = '=RTD("pi.rtdserver", ,'+symbol+', "BidSize")'
sht.range('E2').value = '=RTD("pi.rtdserver", ,'+symbol+', "AskSize")'
sht.range('F2').value = '=RTD("pi.rtdserver", ,'+symbol+', "Bid")'
sht.range('G2').value = '=RTD("pi.rtdserver", ,'+symbol+', "Ask")'
sht.range('H2').value = '=RTD("pi.rtdserver", ,'+symbol+', "lastUpdateTime")'
   

   
import time
i = 1

while(1):    
    x0 = str(sht.range('B2').value)
    x1 = str(sht.range('C2').value)
    x2 = str(sht.range('D2').value)
    x3 = str(sht.range('E2').value)
    x4 = str(sht.range('F2').value)
    x5 = str(sht.range('G2').value)
    x6 = str(sht.range('H2').value)
       

    String1 = x0+x1+x2+x3+x4+x5+x6 

    print(String1)
    y0 = str(sht.range('B2').value)
    y1 = str(sht.range('C2').value)
    y2 = str(sht.range('D2').value)
    y3 = str(sht.range('E2').value)
    y4 = str(sht.range('F2').value)
    y5 = str(sht.range('G2').value)
    y6 = str(sht.range('H2').value)

    String2 = y0+y1+y2+y3+y4+y5+y6 

    if String1==String2:
        continue    

    else:
    
        wrt.range('A'+str(i+1)).value = x0   
        wrt.range('B'+str(i+1)).value = x1   
        wrt.range('C'+str(i+1)).value = x2   
        wrt.range('D'+str(i+1)).value = x3   
        wrt.range('E'+str(i+1)).value = x4   
        wrt.range('F'+str(i+1)).value = x5 
        wrt.range('G'+str(i+1)).value = x6 

        i=i+1
 
#Shukla idea remove duplicate logic









