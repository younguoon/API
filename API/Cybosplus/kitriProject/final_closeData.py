#coding: utf-8
'''
Created on 2016. 5. 17.

@author: Young
''' 
import win32com.client

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
f= open('codeOnly.txt')
codes = f.readlines()

print "code", "," ,
print "date", "," ,
print "price"
for code in codes:
    code = code.strip()
    instStockChart.SetInputValue(0,'A'+str(code))
    instStockChart.SetInputValue(1, '1')
    instStockChart.SetInputValue(2, 20160722)
    instStockChart.SetInputValue(3, 20140101)
    instStockChart.SetInputValue(4, 650)
    instStockChart.SetInputValue(5, [0,5])    
    instStockChart.SetInputValue(6, ord('D'))
    instStockChart.SetInputValue(9, '1')
    
    instStockChart.BlockRequest()
    

    numData = instStockChart.GetHeaderValue(3)

    for i in range(numData):
        valueDate = instStockChart.GetDataValue(0,i)
        stockPrice = instStockChart.GetDataValue(1,i)
        print code.strip(), ",",
        print str(valueDate).strip(), ",",
        print str(stockPrice).strip()
