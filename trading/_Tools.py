
# coding: utf-8

# In[ ]:


def get_code(name): 
    import win32com.client

    instCpCodeMgr = win32com.client.Dispatcsh("CpUtil.CpCodeMgr")
    codeList = instCpCodeMgr.GetStockListByMarket(1)

    temp = {}
    for code in codeList:
        name = instCpCodeMgr.CodeToName(code)
        temp[name] = code
    
    result = temp[name] 
    
    return result

def get_name(code): 
    import win32com.client

    instCpCodeMgr = win32com.client.Dispatcsh("CpUtil.CpCodeMgr")
    codeList = instCpCodeMgr.GetStockListByMarket(1)

    temp = {}
    for code in codeList:
        name = instCpCodeMgr.CodeToName(code)
        temp[code] = name
    
    result = temp[code] 
    
    return result

def get_min(code,time,day):  # 종목, 기간, 오늘, 시점, 분, 시간간격
    start = (datetime.today() - timedelta(day)).strftime("%Y%m%d") 
    instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
    instStockChart.SetInputValue(0, code )
    instStockChart.SetInputValue(1, ord('1'))
    instStockChart.SetInputValue(2, today)
    instStockChart.SetInputValue(3, start)
    # instStockChart.SetInputValue(4, 1000)
    instStockChart.SetInputValue(5, (0,1,5))
    instStockChart.SetInputValue(6, ord(time))  # 'm' : 분, 'T' : 틱
    instStockChart.SetInputValue(7, 1)      # 데이터 주기
    instStockChart.SetInputValue(9, ord('1'))
    instStockChart.SetInputValue(10, 3)

    instStockChart.BlockRequest()

    numData = instStockChart.GetHeaderValue(3)
    numField = instStockChart.GetHeaderValue(1) 

    temp = {}
    for i in range(numData):
        temp[str(instStockChart.GetDataValue(0, i)) +'.'+ str(instStockChart.GetDataValue(1, i))] = [instStockChart.GetDataValue(2, i)]
    temp = pd.DataFrame(temp).transpose()
    temp.index.names = ['time']
    return temp

