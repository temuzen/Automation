
from xlwings import Book, Sheet, Range, Chart
import xlwings as xw
import time


book1 = xw.Book()

sht = xw.books('book1')
#wrt = xw.books('book1').sheets[1]

Symbols = ['','','"NFO_BANKNIFTY18NOVFUT"','"NFO_BANKNIFTY18OCTFUT"','"NFO_BANKNIFTY18SEPFUT"','"NFO_NIFTY18NOVFUT"','"NFO_NIFTY18OCTFUT"','"NFO_NIFTY18SEPFUT"','"NSE_HDFCBANK-EQ"','"NSE_INDUSINDBK-EQ"','"NSE_KOTAKBANK-EQ"','"NSE_AXISBANK-EQ"','"NSE_ICICIBANK-EQ"','"NSE_SBIN-EQ"','"NSE_CANBK-EQ"','"NSE_YESBANK-EQ"','"NSE_BANKBARODA-EQ"','"NSE_FEDERALBNK-EQ"','"NSE_PNB-EQ"','"CDS_USDINR18SEPFUT"','"CDS_USDINR"']

x = len(Symbols)
print(x)

for i in range(2,):

    sht.range('B'+str(i)).value = '=RTD("pi.rtdserver", ,'+Symbols[i]+', "TradingSymbol")'
    sht.range('C'+str(i)).value = '=RTD("pi.rtdserver", ,'+Symbols[i]+', "Last")'
    sht.range('D'+str(i)).value = '=RTD("pi.rtdserver", ,'+Symbols[i]+', "BidSize")'
    sht.range('E'+str(i)).value = '=RTD("pi.rtdserver", ,'+Symbols[i]+', "AskSize")'
    sht.range('F'+str(i)).value = '=RTD("pi.rtdserver", ,'+Symbols[i]+', "Bid")'
    sht.range('G'+str(i)).value = '=RTD("pi.rtdserver", ,'+Symbols[i]+', "Ask")'
    sht.range('H'+str(i)).value = '=RTD("pi.rtdserver", ,'+Symbols[i]+', "lastUpdateTime")'
      
import time
i = 1


'''

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

    String2 = y0+y1+y2+y3+y4+y5+y6+ 

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
 








'''

