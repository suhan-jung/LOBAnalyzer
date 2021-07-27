import win32com.client

class Creon:
    def __init__(self):
        self.stockcur_handlers = {}  # 주식/업종/ELW시세 subscribe event handlers

    def subscribe_stockcur(self, code, cb):
        if not code.startswith('A'):
            code = 'A' + code
        if code in self.stockcur_handlers:
            return
        obj = win32com.client.Dispatch('DsCbo1.StockCur')
        obj.SetInputValue(0, code)
        handler = win32com.client.WithEvents(obj, StockCurEventHandler)
        handler.set_attrs(obj, cb)
        self.stockcur_handlers[code] = obj
        obj.Subscribe()

    def unsubscribe_stockcur(self, code=None):
        lst_code = []
        if code is not None:
            if not code.startswith('A'):
                code = 'A' + code
            if code not in self.stockcur_handlers:
                return
            lst_code.append(code)
        else:
            lst_code = list(self.stockcur_handlers.keys()).copy()
        for code in lst_code:
            obj = self.stockcur_handlers[code]
            obj.Unsubscribe()
            del self.stockcur_handlers[code]


class StockCurEventHandler:
    def set_attrs(self, obj, cb):
        self.obj = obj
        self.cb = cb

    def OnReceived(self):
        item = {
            'code': self.obj.GetHeaderValue(0),
            'name': self.obj.GetHeaderValue(1),
            'diffratio': self.obj.GetHeaderValue(2),
            'timestamp': self.obj.GetHeaderValue(3),
            'price_open': self.obj.GetHeaderValue(4),
            'price_high': self.obj.GetHeaderValue(5),
            'price_low': self.obj.GetHeaderValue(6),
            'bid_sell': self.obj.GetHeaderValue(7),
            'bid_buy': self.obj.GetHeaderValue(8),
            'cum_volume': self.obj.GetHeaderValue(9),  # 주, 거래소지수: 천주
            'cum_trans': self.obj.GetHeaderValue(10),
            'price': self.obj.GetHeaderValue(13),
            'contract_type': self.obj.GetHeaderValue(14),
            'cum_sellamount': self.obj.GetHeaderValue(15),
            'buy_sellamount': self.obj.GetHeaderValue(16),
            'contract_amount': self.obj.GetHeaderValue(17),
            'second': self.obj.GetHeaderValue(18),
            'price_type': chr(self.obj.GetHeaderValue(19)),  # 1: 동시호가시간 예상체결가, 2: 장중 체결가
            'market_flag': chr(self.obj.GetHeaderValue(20)),  # '1': 장전예상체결, '2': 장중, '4': 장후시간외, '5': 장후예상체결
            'premarket_volume': self.obj.GetHeaderValue(21),
            'diffsign': chr(self.obj.GetHeaderValue(22)),
            'LP보유수량':self.obj.GetHeaderValue(23),
            'LP보유수량대비':self.obj.GetHeaderValue(24),
            'LP보유율':self.obj.GetHeaderValue(25),
            '체결상태(호가방식)':self.obj.GetHeaderValue(26),
            '누적매도체결수량(호가방식)':self.obj.GetHeaderValue(27),
            '누적매수체결수량(호가방식)':self.obj.GetHeaderValue(28),
        }
        self.cb(item)