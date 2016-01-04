import sys
import win32com
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from win32com import client 
#from quamash import QEventLoop

qt5_url = "C:/cygwin/home/utylee/.virtualenvs/tyTrader-win/Lib/site-packages/PyQt5/plugins" 

class MyWindows(QMainWindow):
#class MyWindows:
    def __init__(self):
        super().__init__()

        self.setWindowTitle("test")

        #test =win32com.client.Dispatch("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        print('테스트 입니당 데헷')
        self.kiwoom.dynamicCall("CommConnect()")


if __name__ == "__main__":
    print('갑니다')
    #QtCore.QCoreApplication.setLibraryPaths(['C:/cygwin/home/utylee/.virtualenvs/tyTrader-win/Lib/site-packages/PyQt5/plugins'])
    QCoreApplication.setLibraryPaths([qt5_url])

    app = QApplication(sys.argv)
    window = MyWindows()
    window.show()
    app.exec_()
#test.CommConnect()
