import sys
import asyncio
import win32com
#from PyQt5.QtGui import QGuiApplication
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from win32com import client 
from quamash import QEventLoop

qt5_url = "C:/cygwin/home/utylee/.virtualenvs/tyTrader-win/Lib/site-packages/PyQt5/plugins" 

# PyQt와 quamash를 이용해서 asyncio 방식을 사용합니다
# PyQt5를 virtualenv 상에서 사용하기 위해서는 정확하게 Platform 폴더를 지정해줘야 한다고 합니다.
QCoreApplication.setLibraryPaths([qt5_url])
app = QApplication(sys.argv)
loop = QEventLoop(app)
asyncio.set_event_loop(loop)

class MyWindows(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("test")

        #test =win32com.client.Dispatch("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.dynamicCall("CommConnect()")

with loop:
    pass
    window = MyWindows()
    window.show()
    loop.run_forever()

'''
if __name__ == "__main__":
    QCoreApplication.setLibraryPaths([qt5_url])

    #app = QApplication(sys.argv)
    window = MyWindows()
    window.show()
    #app.exec_()
'''
