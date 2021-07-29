from creon import Creon
from PyQt5.QtWidgets import *
import sys
from csv import DictWriter
import datetime
import os.path

c = Creon()
date_today = datetime.datetime.today().strftime('%Y%m%d')

def cb_futures(item):
    # print(item)
    headersCSV = [ #컬럼순서 조정필요
        'code',
        'trade_quote',
        'time',
        'time_actual',
        'trade_price',
        'open',
        'high',
        'low',
        'ask_price_trade',
        'bid_price_trade',
        'ask_amount_trade',
        'bid_amount_trade',
        'change',
        'volume',
        'acc_volume',
        'acc_buy_volume',
        'acc_sell_volume',
        'trade_buy_sell',
        'trade_incoming_option',
        'ask_price_1',
        'ask_price_2',
        'ask_price_3',
        'ask_price_4',
        'ask_price_5',
        'ask_amount_1',
        'ask_amount_2',
        'ask_amount_3',
        'ask_amount_4',
        'ask_amount_5',
        'ask_amount_total',
        'ask_count_1',
        'ask_count_2',
        'ask_count_3',
        'ask_count_4',
        'ask_count_5',
        'ask_count_total',
        'bid_price_1',
        'bid_price_2',
        'bid_price_3',
        'bid_price_4',
        'bid_price_5',
        'bid_amount_1',
        'bid_amount_2',
        'bid_amount_3',
        'bid_amount_4',
        'bid_amount_5',
        'bid_amount_total',
        'bid_count_1',
        'bid_count_2',
        'bid_count_3',
        'bid_count_4',
        'bid_count_5',
        'bid_count_total',
    ]
    code = item['code']
    filename = f'./csv/futures_{code}_{date_today}.csv'
    file_exists = os.path.isfile(filename)
    with open(filename, 'a', newline='') as f_object:
        # Pass the CSV  file object to the Dictwriter() function
        # Result - a DictWriter object
        dictwriter_object = DictWriter(f_object, fieldnames=headersCSV)
        # Pass the data in the dictionary as an argument into the writerow() function
        if not file_exists:
            dictwriter_object.writeheader()
        dictwriter_object.writerow(item)
        # Close the file object
        f_object.close()
        # print(item['code'], item['trade_price'])
        # item을 처리하는 모듈 작성

def cb_stock(item):
    # print(item)
    headersCSV = [
        'code',
        'trade_quote',
        'time',
        'time_actual',
        'trade_price',
        'change',
        'open',
        'high',
        'low',
        'ask_price_trade',
        'bid_price_trade',
        'acc_volume',
        'acc_amount',
        'trade_buy_sell',
        'acc_buy_volume',
        'acc_sell_volume',
        'volume',
        'second',
        'price_type',
        'market_flag',
        'premarket_volume',
        'diffsign',
        'LP보유수량',
        'LP보유수량대비',
        'LP보유율',
        '체결상태(호가방식)',
        '누적매도체결수량(호가방식)',
        '누적매수체결수량(호가방식)',
        'ask_price_1',
        'ask_price_2',
        'ask_price_3',
        'ask_price_4',
        'ask_price_5',
        'ask_amount_1',
        'ask_amount_2',
        'ask_amount_3',
        'ask_amount_4',
        'ask_amount_5',
        'ask_amount_total',
        'ask_count_1',
        'ask_count_2',
        'ask_count_3',
        'ask_count_4',
        'ask_count_5',
        'ask_count_total',
        'bid_price_1',
        'bid_price_2',
        'bid_price_3',
        'bid_price_4',
        'bid_price_5',
        'bid_amount_1',
        'bid_amount_2',
        'bid_amount_3',
        'bid_amount_4',
        'bid_amount_5',
        'bid_amount_total',
        'bid_count_1',
        'bid_count_2',
        'bid_count_3',
        'bid_count_4',
        'bid_count_5',
        'bid_count_total',
        
        'ask_price_1',
        'ask_price_2',
        'ask_price_3',
        'ask_price_4',
        'ask_price_5',
        'ask_amount_1',
        'ask_amount_2',
        'ask_amount_3',
        'ask_amount_4',
        'ask_amount_5',
        'ask_amount_total',
        'ask_count_1',
        'ask_count_2',
        'ask_count_3',
        'ask_count_4',
        'ask_count_5',
        'ask_count_total',
        'bid_price_1',
        'bid_price_2',
        'bid_price_3',
        'bid_price_4',
        'bid_price_5',
        'bid_amount_1',
        'bid_amount_2',
        'bid_amount_3',
        'bid_amount_4',
        'bid_amount_5',
        'bid_amount_total',
        'bid_count_1',
        'bid_count_2',
        'bid_count_3',
        'bid_count_4',
        'bid_count_5',
        'bid_count_total',
    ]
    code = item['code']
    filename = f'./csv/futures_{code}_{date_today}.csv'
    file_exists = os.path.isfile(filename)
    with open(filename, 'a', newline='') as f_object:
        # Pass the CSV  file object to the Dictwriter() function
        # Result - a DictWriter object
        dictwriter_object = DictWriter(f_object, fieldnames=headersCSV)
        # Pass the data in the dictionary as an argument into the writerow() function
        if not file_exists:
            dictwriter_object.writeheader()
        dictwriter_object.writerow(item)
        # Close the file object
        f_object.close()
        # print(item['code'], item['trade_price'])
        # item을 처리하는 모듈 작성
# c.subscribe_stockcur('006800', cb)
# # ...
# c.unsubscribe_stockcur()

class MyWindow(QMainWindow):
 
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 180)
        self.dicCurData = {}
 
        btnStart = QPushButton("요청 시작", self)
        btnStart.move(20, 20)
        btnStart.clicked.connect(self.btnStart_clicked)
 
        btnStop = QPushButton("요청 종료", self)
        btnStop.move(20, 70)
        btnStop.clicked.connect(self.btnStop_clicked)
 
        btnExit = QPushButton("종료", self)
        btnExit.move(20, 120)
        btnExit.clicked.connect(self.btnExit_clicked)
 
 
 
    def btnStart_clicked(self):
        c.subscribe_futures('167R9', cb_futures)
        c.subscribe_futures('165R9', cb_futures)
        # c.subscribe_stocktrade('035420', cb)
        # c.subscribe_stocktrade('025980', cb)
        # objCur = CpRPOvForMst()
        # objCur.Request('167R9', self.dicCurData)
 
 
    def btnStop_clicked(self):
        c.unsubscribe_stock()
        c.unsubscribe_futures()
        return
 
 
    def btnExit_clicked(self):
        exit()
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
