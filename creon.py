import win32com.client
import datetime

class Creon:
    def __init__(self):
        self.stock_handlers = {}  # 주식/업종/ELW시세 subscribe event handlers
        self.futures_handlers = {}  # futures subscribe event handlers

    def subscribe_stock(self, code, cb):
        if not code.startswith('A'):
            code = 'A' + code
        if code in self.stock_handlers:
            return
        obj_trade = win32com.client.Dispatch('DsCbo1.StockCur')
        obj_quote = win32com.client.Dispatch('DsCbo1.StockJpBid')
        obj_trade.SetInputValue(0, code)
        obj_quote.SetInputValue(0, code)
        handler_trade = win32com.client.WithEvents(obj_trade, StockEventHandler)
        handler_quote = win32com.client.WithEvents(obj_quote, StockEventHandler)
        handler_trade.set_attrs(obj_trade, cb, 'trade')
        handler_quote.set_attrs(obj_quote, cb, 'quote')
        self.stock_handlers[code] = [obj_trade,obj_quote]
        obj_trade.Subscribe()
        obj_quote.Subscribe()

    def unsubscribe_stock(self, code=None):
        lst_code = []
        if code is not None:
            if not code.startswith('A'):
                code = 'A' + code
            if code not in self.stock_handlers:
                return
            lst_code.append(code)
        else:
            lst_code = list(self.stock_handlers.keys()).copy()
        for code in lst_code:
            for obj in self.stock_handlers[code]:
                obj.Unsubscribe()
            del self.stock_handlers[code]

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
                'time_received' : datetime.datetime.now().strftime('%H:%M:%S.%f'),
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
                'time_received' : datetime.datetime.now().strftime('%H:%M:%S.%f'),
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
        
class StockEventHandler:
    def set_attrs(self, obj, cb, name):
        self.obj = obj
        self.cb = cb
        self.name = name

    def OnReceived(self):
        if self.name == 'trade':
            item = {
                'code': self.obj.GetHeaderValue(0),
                'name': self.obj.GetHeaderValue(1),
                'trade_quote' : 'trade',
                'time_received' : datetime.datetime.now().strftime('%H:%M:%S.%f'),
                'time': self.obj.GetHeaderValue(3),
                'trade_price': self.obj.GetHeaderValue(13),
                'change': int(self.obj.GetHeaderValue(2)),
                'open': self.obj.GetHeaderValue(4),
                'high': self.obj.GetHeaderValue(5),
                'low': self.obj.GetHeaderValue(6),
                'ask_price_trade': self.obj.GetHeaderValue(7),
                'bid_price_trade': self.obj.GetHeaderValue(8),
                'acc_volume': self.obj.GetHeaderValue(9),  # 주, 거래소지수: 천주
                'acc_amount': self.obj.GetHeaderValue(10), # 거래대금
                'trade_buy_sell': chr(self.obj.GetHeaderValue(14)),
                'acc_sell_volume': self.obj.GetHeaderValue(15),
                'acc_buy_volume': self.obj.GetHeaderValue(16),
                'volume': self.obj.GetHeaderValue(17),
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
        else:
            item = {
                'code': self.obj.GetHeaderValue(0),
                'trade_quote' : 'quote',
                'time_received' : datetime.datetime.now().strftime('%H:%M:%S.%f'),
                'time' : self.obj.GetHeaderValue(1),  # 처리시각
                'ask_price_1' : int(self.obj.GetHeaderValue(3)),
                'bid_price_1' : int(self.obj.GetHeaderValue(4)),
                'ask_amount_1' : int(self.obj.GetHeaderValue(5)),
                'bid_amount_1' : int(self.obj.GetHeaderValue(6)),
                'ask_price_2' : int(self.obj.GetHeaderValue(7)),
                'bid_price_2' : int(self.obj.GetHeaderValue(8)),
                'ask_amount_2' : int(self.obj.GetHeaderValue(9)),
                'bid_amount_2' : int(self.obj.GetHeaderValue(10)),
                'ask_price_3' : int(self.obj.GetHeaderValue(11)),
                'bid_price_3' : int(self.obj.GetHeaderValue(12)),
                'ask_amount_3' : int(self.obj.GetHeaderValue(13)),
                'bid_amount_3' : int(self.obj.GetHeaderValue(14)),
                'ask_price_4' : int(self.obj.GetHeaderValue(15)),
                'bid_price_4' : int(self.obj.GetHeaderValue(16)),
                'ask_amount_4' : int(self.obj.GetHeaderValue(17)),
                'bid_amount_4' : int(self.obj.GetHeaderValue(18)),
                'ask_price_5' : int(self.obj.GetHeaderValue(19)),
                'bid_price_5' : int(self.obj.GetHeaderValue(20)),
                'ask_amount_5' : int(self.obj.GetHeaderValue(21)),
                'bid_amount_5' : int(self.obj.GetHeaderValue(22)),
                'ask_amount_total' : int(self.obj.GetHeaderValue(23)),
                'bid_amount_total' : int(self.obj.GetHeaderValue(24)),
                'ask_amount_total_off' : int(self.obj.GetHeaderValue(25)),
                'bid_amount_total_off' : int(self.obj.GetHeaderValue(26)),
                'ask_price_6' : int(self.obj.GetHeaderValue(27)),
                'bid_price_6' : int(self.obj.GetHeaderValue(28)),
                'ask_amount_6' : int(self.obj.GetHeaderValue(29)),
                'bid_amount_6' : int(self.obj.GetHeaderValue(30)),
                'ask_price_7' : int(self.obj.GetHeaderValue(31)),
                'bid_price_7' : int(self.obj.GetHeaderValue(32)),
                'ask_amount_7' : int(self.obj.GetHeaderValue(33)),
                'bid_amount_7' : int(self.obj.GetHeaderValue(34)),
                'ask_price_8' : int(self.obj.GetHeaderValue(35)),
                'bid_price_8' : int(self.obj.GetHeaderValue(36)),
                'ask_amount_8' : int(self.obj.GetHeaderValue(37)),
                'bid_amount_8' : int(self.obj.GetHeaderValue(38)),
                'ask_price_9' : int(self.obj.GetHeaderValue(39)),
                'bid_price_9' : int(self.obj.GetHeaderValue(40)),
                'ask_amount_9' : int(self.obj.GetHeaderValue(41)),
                'bid_amount_9' : int(self.obj.GetHeaderValue(42)),
                'ask_price_10' : int(self.obj.GetHeaderValue(43)),
                'bid_price_10' : int(self.obj.GetHeaderValue(44)),
                'ask_amount_10' : int(self.obj.GetHeaderValue(45)),
                'bid_amount_10' : int(self.obj.GetHeaderValue(46)),
                'ask_amount_LP_1' : int(self.obj.GetHeaderValue(47)),
                'bid_amount_LP_1' : int(self.obj.GetHeaderValue(48)),
                'ask_amount_LP_2' : int(self.obj.GetHeaderValue(49)),
                'bid_amount_LP_2' : int(self.obj.GetHeaderValue(50)),
                'ask_amount_LP_3' : int(self.obj.GetHeaderValue(51)),
                'bid_amount_LP_3' : int(self.obj.GetHeaderValue(52)),
                'ask_amount_LP_4' : int(self.obj.GetHeaderValue(53)),
                'bid_amount_LP_4' : int(self.obj.GetHeaderValue(54)),
                'ask_amount_LP_5' : int(self.obj.GetHeaderValue(55)),
                'bid_amount_LP_5' : int(self.obj.GetHeaderValue(56)),
                'ask_amount_LP_6' : int(self.obj.GetHeaderValue(57)),
                'bid_amount_LP_6' : int(self.obj.GetHeaderValue(58)),
                'ask_amount_LP_7' : int(self.obj.GetHeaderValue(59)),
                'bid_amount_LP_7' : int(self.obj.GetHeaderValue(60)),
                'ask_amount_LP_8' : int(self.obj.GetHeaderValue(61)),
                'bid_amount_LP_8' : int(self.obj.GetHeaderValue(62)),
                'ask_amount_LP_9' : int(self.obj.GetHeaderValue(63)),
                'bid_amount_LP_9' : int(self.obj.GetHeaderValue(64)),
                'ask_amount_LP_10' : int(self.obj.GetHeaderValue(65)),
                'bid_amount_LP_10' : int(self.obj.GetHeaderValue(66)),
                'ask_amount_LP_total' : int(self.obj.GetHeaderValue(67)),
                'bid_amount_LP_total' : int(self.obj.GetHeaderValue(68)),
            }
        self.cb(item)