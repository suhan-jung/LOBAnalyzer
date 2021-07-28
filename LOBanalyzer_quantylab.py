from creon import Creon
from PyQt5.QtWidgets import *
import sys
from csv import DictWriter
import datetime

c = Creon()
date_today = datetime.datetime.today().strftime('%Y%m%d')
headersCSV = [
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
    ]

def cb(item):
    print(item)
    with open(f'./csv/futures_{date_today}.csv', 'a', newline='') as f_object:
        # Pass the CSV  file object to the Dictwriter() function
        # Result - a DictWriter object
        dictwriter_object = DictWriter(f_object, fieldnames=headersCSV)
        # Pass the data in the dictionary as an argument into the writerow() function
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
        c.subscribe_futures('167R9', cb)
        # c.subscribe_stocktrade('035420', cb)
        # c.subscribe_stocktrade('025980', cb)
        # objCur = CpRPOvForMst()
        # objCur.Request('167R9', self.dicCurData)
 
 
    def btnStop_clicked(self):
        c.unsubscribe_stocktrade()
        c.unsubscribe_futures()
        return
 
 
    def btnExit_clicked(self):
        exit()
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
