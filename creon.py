import win32com.client
import datetime

class Creon:
    def __init__(self):
        self.stocktrade_handlers = {}  # 주식/업종/ELW시세 subscribe event handlers
        self.futures_handlers = {}  # futures subscribe event handlers

    def subscribe_stocktrade(self, code, cb):
        if not code.startswith('A'):
            code = 'A' + code
        if code in self.stocktrade_handlers:
            return
        obj = win32com.client.Dispatch('DsCbo1.StockCur')
        obj.SetInputValue(0, code)
        handler = win32com.client.WithEvents(obj, StockTradeEventHandler)
        handler.set_attrs(obj, cb)
        self.stocktrade_handlers[code] = obj
        obj.Subscribe()

    def unsubscribe_stocktrade(self, code=None):
        lst_code = []
        if code is not None:
            if not code.startswith('A'):
                code = 'A' + code
            if code not in self.stocktrade_handlers:
                return
            lst_code.append(code)
        else:
            lst_code = list(self.stocktrade_handlers.keys()).copy()
        for code in lst_code:
            obj = self.stocktrade_handlers[code]
            obj.Unsubscribe()
            del self.stocktrade_handlers[code]

    def subscribe_futures(self, code, cb):
        if code in self.futures_handlers:
            return
        obj_trade = win32com.client.Dispatch('DsCbo1.FutureCurOnly')
        obj_quote = win32com.client.Dispatch('CpSysDib.FutureJpBid')
        obj_trade.SetInputValue(0, code)
        obj_quote.SetInputValue(0, code)
        handler_trade = win32com.client.WithEvents(obj_trade, FuturesEventHandler)
        handler_quote = win32com.client.WithEvents(obj_quote, FuturesEventHandler)
        handler_trade.set_attrs(obj_trade, cb, "trade")
        handler_quote.set_attrs(obj_quote, cb, "quote")
        self.futures_handlers[code] = [obj_trade, obj_quote]
        obj_trade.Subscribe()
        obj_quote.Subscribe()
            
    def unsubscribe_futures(self, code=None):
        lst_code = []
        if code is not None:
            if code not in self.futures_handlers:
                return
            lst_code.append(code)
        else:
            lst_code = list(self.futures_handlers.keys()).copy()
        for code in lst_code:
            for obj in self.futures_handlers[code]:
                obj.Unsubscribe()
            del self.futures_handlers[code]
            
class FuturesEventHandler:
    def set_attrs(self, obj, cb, name):
        self.obj = obj
        self.cb = cb
        self.name = name

    def OnReceived(self):
        if self.name == 'trade':
            item = {
                'code': self.obj.GetHeaderValue(0),
                'trade_quote' : 'trade',
                'time_actual' : datetime.datetime.now().strftime('%H:%M:%S.%f'),
                'trade_price' : round(self.obj.GetHeaderValue(1),2), # 현재가
                'open' : self.obj.GetHeaderValue(7), # 시가
                'high' : self.obj.GetHeaderValue(8), # 고가
                'low' : self.obj.GetHeaderValue(9), # 저가
                'ask_price_trade' : round(self.obj.GetHeaderValue(18),2),  # 최우선매도호가
                'bid_price_trade' : round(self.obj.GetHeaderValue(19),2),  # 최우선매수호가
                'ask_amount_trade' : self.obj.GetHeaderValue(20),  # 최우선매도수량
                'bid_amount_trade' : self.obj.GetHeaderValue(21),  # 최우선매수수량
                'change' : self.obj.GetHeaderValue(2),  # 전일대비
                'volume' : self.obj.GetHeaderValue(13),  # 거래량
                'acc_volume' : self.obj.GetHeaderValue(13),  # 누적거래량
                'acc_buy_volume' : self.obj.GetHeaderValue(23),  # 누적매수체결
                'acc_sell_volume' : self.obj.GetHeaderValue(22),  # 누적매도체결
                'time' : self.obj.GetHeaderValue(15),  # 체결시각
                'trade_buy_sell' : chr(self.obj.GetHeaderValue(24)),  # 체결구분 1매수2매도
                'trade_incoming_option' : chr(self.obj.GetHeaderValue(30)),  # 체결수신구분 1체결 2체결+호가잔량
            }
        else:
            item = {
                'code': self.obj.GetHeaderValue(0),
                'trade_quote' : 'quote',
                'time_actual' : datetime.datetime.now().strftime('%H:%M:%S.%f'),
                'time' : self.obj.GetHeaderValue(1),  # 처리시각
                'ask_price_1' : round(self.obj.GetHeaderValue(2),2),
                'ask_price_2' : round(self.obj.GetHeaderValue(3),2),
                'ask_price_3' : round(self.obj.GetHeaderValue(4),2),
                'ask_price_4' : round(self.obj.GetHeaderValue(5),2),
                'ask_price_5' : round(self.obj.GetHeaderValue(6),2),
                'ask_amount_1' : int(self.obj.GetHeaderValue(7)),
                'ask_amount_2' : int(self.obj.GetHeaderValue(8)),
                'ask_amount_3' : int(self.obj.GetHeaderValue(9)),
                'ask_amount_4' : int(self.obj.GetHeaderValue(10)),
                'ask_amount_5' : int(self.obj.GetHeaderValue(11)),
                'ask_amount_total' : int(self.obj.GetHeaderValue(12)),
                'ask_count_1' : int(self.obj.GetHeaderValue(13)),
                'ask_count_2' : int(self.obj.GetHeaderValue(14)),
                'ask_count_3' : int(self.obj.GetHeaderValue(15)),
                'ask_count_4' : int(self.obj.GetHeaderValue(16)),
                'ask_count_5' : int(self.obj.GetHeaderValue(17)),
                'ask_count_total' : int(self.obj.GetHeaderValue(18)),
                'bid_price_1' : round(self.obj.GetHeaderValue(19),2),
                'bid_price_2' : round(self.obj.GetHeaderValue(20),2),
                'bid_price_3' : round(self.obj.GetHeaderValue(21),2),
                'bid_price_4' : round(self.obj.GetHeaderValue(22),2),
                'bid_price_5' : round(self.obj.GetHeaderValue(23),2),
                'bid_amount_1' : int(self.obj.GetHeaderValue(24)),
                'bid_amount_2' : int(self.obj.GetHeaderValue(25)),
                'bid_amount_3' : int(self.obj.GetHeaderValue(26)),
                'bid_amount_4' : int(self.obj.GetHeaderValue(27)),
                'bid_amount_5' : int(self.obj.GetHeaderValue(28)),
                'bid_amount_total' : int(self.obj.GetHeaderValue(29)),
                'bid_count_1' : int(self.obj.GetHeaderValue(30)),
                'bid_count_2' : int(self.obj.GetHeaderValue(31)),
                'bid_count_3' : int(self.obj.GetHeaderValue(32)),
                'bid_count_4' : int(self.obj.GetHeaderValue(33)),
                'bid_count_5' : int(self.obj.GetHeaderValue(34)),
                'bid_count_total' : int(self.obj.GetHeaderValue(35)),
            }
        self.cb(item)
        
class StockTradeEventHandler:
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