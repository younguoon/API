# -*- coding: utf-8 -*-
import win32com.client

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")

codeList = instCpCodeMgr.GetStockListByMarket(1)
codeList = codeList + instCpCodeMgr.GetStockListByMarket(2)

instMarketEye.SetInputValue(0,[20])

print "code", "," ,
print "name", "," ,
print "totalStrock"

for i, code in enumerate(codeList):
    secondCode = instCpCodeMgr.GetStockSectionKind(code)
    name = instCpCodeMgr.CodeToName(code)
    marketCode = instCpCodeMgr.GetStockMarketKind(code)

    instMarketEye.SetInputValue(1, [code])
    instMarketEye.BlockRequest()
    totalStock = instMarketEye.GetDataValue(0,0)

    if len(name)!=0 and secondCode==1:
    	#print i, "," , code ,"," , secondCode,",", name, ",", marketCode
    	#print i 	, "," ,
    	print code[1:]	, "," , #218150, secondCode,",", name, ",", marketCode
    	print name.encode('utf-8') , "," ,
    	print totalStock

