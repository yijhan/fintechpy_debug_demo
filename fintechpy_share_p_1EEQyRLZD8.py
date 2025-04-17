# 相關討論串
# https://www.facebook.com/share/p/1EEQyRLZD8/

import pythoncom, time
import comtypes.client as cc
cc.GetModule(r'C:/SKCOM/SKCOM.dll')
import comtypes.gen.SKCOMLib as sk
import asyncio
import math

ts = sk.SKSTOCK()
skC = cc.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
skQ = cc.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
skR = cc.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)
skO = cc.CreateObject(sk.SKOrderLib, interface=sk.ISKOrderLib)

# 登入帳號密碼
ID = ""
PW = ""
Account = ""

# 想取得報價的股票代碼
strStocks = 'TSEA'
#strStocks ='2317'

# working functions, async coruntime to pump events
async def pump_task():
    while True:
        pythoncom.PumpWaitingMessages()
        await asyncio.sleep(0.1)

def message_pump(duration_seconds=1):
    """
    Pump COM messages for specified duration
    Args:
        duration_seconds (int): How long to pump messages
    """
    start_time = time.time()
    while time.time() - start_time < duration_seconds:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.01)  # Prevent CPU overload

class skR_events:
    def OnReplyMessage(self, bstrUserID, bstrMessage):
        sConfirmCode = -1
        print('OnReplyMessage', bstrUserID, bstrMessage)
        return sConfirmCode

    def OnNewData(self, bstrUserID, bstrData):
        print('skR_OnNewData', bstrData)

# 建立事件類別
class skQ_events:
    def OnConnection(self, nKind, nCode):
        if nCode == 0:
            if nKind == 3001:
                print("skQ OnConnection, nkind= ", nKind)
            elif nKind == 3003:
                # 等到回報3003 確定連線報價伺服器成功後，才登陸要報價的股票
                skQ.SKQuoteLib_RequestStocks(1, strStocks)
                print("skQ OnConnection, request stocks, nkind= ", nKind)

    def OnNotifyQuoteLONG(self, sMarketNo, sStockIdx):
        pStock = sk.SKSTOCKLONG()
        m_ncode = skQ.SKQuoteLib_GetStockByIndexLONG(sMarketNo, sStockIdx, pStock)
        strMsg = ('代碼:', pStock.bstrStockNo, 
                 '--名稱:', pStock.bstrStockName, 
                 '--開盤價:', pStock.nOpen / math.pow(10, pStock.sDecimal),
                 '--最高:', pStock.nHigh / math.pow(10, pStock.sDecimal),
                 '--最低:', pStock.nLow / math.pow(10, pStock.sDecimal),
                 '--成交價:', pStock.nClose / math.pow(10, pStock.sDecimal),
                 '--總量:', pStock.nTQty)
        if len(strMsg) != 0:
            print(strMsg)

class skO_events():
    def OnRealBalanceReport(self, bstrData):
        msg = bstrData.split(',')
        # class = 股票種類 , T=集保, C=融資, L=融券
        dic = {'stk_no': msg[0], 'class': msg[1], 'instock': msg[6]}
        print(bstrData)
        print(dic)
        pass

# Event sink
EventQ = skQ_events()
EventR = skR_events()
EventO = skO_events()

# make connection to event sink
ConnectionQ = cc.GetEvents(skQ, EventQ)
ConnectionR = cc.GetEvents(skR, EventR)
ConnectionO = cc.GetEvents(skO, EventO)

# get an event loop
pumping_loop = asyncio.get_event_loop().create_task(pump_task())
print('Event pumping!')

# Login

print("Login,", skC.SKCenterLib_GetReturnCodeMessage(skC.SKCenterLib_Login(ID, PW)))
message_pump(1)

print("SKReplyLib_ConnectByID,", skC.SKCenterLib_GetReturnCodeMessage(skR.SKReplyLib_ConnectByID(ID)))
message_pump(1)

# 初始化 SKOrderLib
ncode = skO.SKOrderLib_Initialize()
print("SKOrderLib_Initialize", skC.SKCenterLib_GetReturnCodeMessage(ncode))


# 登錄報價伺服器，範例在OnConnection收到3003，會順便註冊報價股票
print("EnterMonitor,", skC.SKCenterLib_GetReturnCodeMessage(skQ.SKQuoteLib_EnterMonitorLONG()))
# 每秒 pump event 一次，這裡示範10秒
message_pump(10)

# 離線後都要重新連線，再重新註冊股票才會報價。
# m_ncode = skQ.SKQuoteLib_LeaveMonitor()
# print("Leave Monitor...", m_ncode)

# 切換股票，需先RequestStocks，也只有在收到3003後才能RequestStocks
strStocks = '2317'
page, ncode = skQ.SKQuoteLib_RequestStocks(1, strStocks) # RequestStocks 會回傳兩個參數
print(f"RequestStocks {strStocks}", skC.SKCenterLib_GetReturnCodeMessage(ncode))
message_pump(6)


m_nCode = skO.GetRealBalanceReport(ID, Account)
print('GetRealBalanceReport', m_nCode)

message_pump(6)
