from creon import Creon
from PyQt5.QtWidgets import *
import sys

c = Creon()

def cb(item):
    print(item['name'], item['price'])


c.subscribe_stockcur('006800', cb)

# ...

c.unsubscribe_stockcur()


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
        c.subscribe_stockcur('035420', cb)
        # objCur = CpRPOvForMst()
        # objCur.Request('167R9', self.dicCurData)
 
 
    def btnStop_clicked(self):
        return
 
 
    def btnExit_clicked(self):
        exit()
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
