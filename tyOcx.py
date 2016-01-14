import sys
import asyncio
import time
import socket
import win32com

#from PyQt5.QtGui import QGuiApplication
from PyQt5.QtGui import *
from PyQt5.QtCore import pyqtSlot, QCoreApplication
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
#from win32com import client 
from quamash import QEventLoop

from tyLogic import *
#from tyUtils import *
import tyUtils

qt5_url = "C:/cygwin/home/utylee/.virtualenvs/tyTrader-win/Lib/site-packages/PyQt5/plugins" 

# PyQt5를 virtualenv 상에서 사용하기 위해서는 정확하게 Platform 폴더를 지정해줘야 한다고 합니다.
QCoreApplication.setLibraryPaths([qt5_url])
# PyQt와 quamash를 이용해서 asyncio eventloop방식을 사용합니다
app = QApplication(sys.argv)
loop = QEventLoop(app)
asyncio.set_event_loop(loop)

# 키움 ocx를 글로벌 static 변수로 밖으로 빼놓았습니다.
#kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

# 인터페이스 윈도우입니다.
class MyWindow(QMainWindow):
    def __init__(self, loop):
        super().__init__()

        self.elapsed = 0
        self.loop = loop

        self.setWindowTitle("tyOcx")
        self.setGeometry(1800,300,600,700)

        self.ordered = 0   # 주문 발생 여부, 중복 주문을 막기 위한 변수입니다
        self.start_timer = 0    # 타이머 시작명령(실은 미리 시작했습니다. 다시 말해 적용을 알리는) 플래그


        #test =win32com.client.Dispatch("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

    #@asyncio.coroutine
    def initialize(self):

        # 종목 세트를 선언하고 ocx를 설정해줍니다.
        self.jongmok_set = Jongmok_set(self.kiwoom)

        # 접속 버튼을 생성합니다
        self.btn1 = QPushButton("Acc Info", self)
        self.btn1.move(20, 20)
        self.btn1.clicked.connect(self.OnBtn1_clicked)
        #self.connect(self.btn1, SIGNAL("clicked()"), self.OnBtn1_clicked)

        # 각종 ocx receive 시그널들을 연결해줍니다.
        self.kiwoom.OnReceiveTrData.connect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveRealData.connect(self.OnReceiveRealData)
        self.kiwoom.OnReceiveChejanData.connect(self.OnReceiveChejanData)
        self.kiwoom.OnReceiveMsg.connect(self.OnReceiveMsg)


        # 주가 정보 요청 버튼을 생성합니다
        self.btn2 = QPushButton("request", self)
        self.btn2.move(20, 70)
        self.btn2.clicked.connect(self.OnBtn2_clicked)
        #self.connect(self.btn2, SIGNAL("clicked()"), self.OnBtn2_clicked)

        # 실시간 시세 추가 버튼을 생성합니다.
        self.btn3 = QPushButton("realtime", self)
        self.btn3.move(20, 100)
        self.btn3.clicked.connect(self.OnBtn3_clicked)

        # 매수 주문 버튼을 생성합니다.
        self.btn4 = QPushButton("order", self)
        self.btn4.move(20, 130)
        self.btn4.clicked.connect(self.OnBtn4_clicked)

        # 조건 매수 버튼을 생성합니다.
        self.btn5 = QPushButton("add condi", self)
        self.btn5.setGeometry(20, 170, 200, 50)
        #self.btn5.move(20, 170)
        self.btn5.clicked.connect(self.OnBtn5_clicked)

        # 분봉 조회 버튼을 생성합니다.
        self.btn6 = QPushButton("min candle", self)
        self.btn6.setGeometry(20, 250, 150, 40)
        self.btn6.clicked.connect(self.OnBtn6_clicked)

        # 키움OpenApi 접속 창을 띄웁니다.
        #self.kiwoom.dynamicCall("CommConnect()")
        print(".접속중입니다. 기다려주십시오...")

        print('calling connecting')
        #asyncio.async(self.connect_and_timer())
        self.loop.run_until_complete(self.connect_and_timer())
        #self.connecting()

        # 타이머를 시작합니다.
        #yield from asyncio.async(self.timer_async())

    @asyncio.coroutine
    def connect_and_timer(self):
        ret = self.kiwoom.CommConnect()
        yield from asyncio.sleep(6)
        print('after 6 seconds..')
        print('starting asyncio timer...')
        asyncio.async(self.timer_async())

        #self.loop.call_later(3, self.timer_async)

    def connect_kiwoom(self):
        #ret = self.kiwoom.dynamicCall("CommConnect()")
        print('come into connect_kiwoom')
        ret = self.kiwoom.CommConnect()

    def OnBtn1_clicked(self):
        #ret = self.kiwoom.dynamicCall("CommConnect()")
        #ret = self.kiwoom.CommConnect()

        ret = self.kiwoom.GetLoginInfo("ACCNO")
        print(ret.strip())

    def OnBtn2_clicked(self):
        print('request emitted')
        #ret = self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", "053260") # 금강철강
        #ret = self.kiwoom.SetInputValue("종목코드", "016090")  # 대현
        ret = self.kiwoom.SetInputValue("종목코드", "032540")  # TJ미디어
        #ret = self.kiwoom.SetInputValue(QString("종목코드"), QString("053260"))
        #ret = self.kiwoom.SetInputValue("종목코드", "016090")
        print(ret)
        #ret = self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "주식기본정보",\
        #                            "OPT10001", 0, "0101")
        ret = self.kiwoom.CommRqData("주식기본정보", "OPT10001", 0, "0101")
        print(ret)

    # real time 실시간 버튼이 눌렸을 때의 로직함수입니다
    def OnBtn3_clicked(self):
        #ret = self.kiwoom.dynamicCall("SetRealReg(QString, QString, int, QString)", "9999", "053260", 0, "0")
        #ret = self.kiwoom.dynamicCall("SetRealReg(QString, QString, int, QString)", "9999", "016090", 0, "0")
        #ret = self.kiwoom.SetRealReg("0001", "032540", "10", "0") #TJ 미디어


        ret = self.kiwoom.SetRealReg("0001", "053260", "10", "0") #금강철강
        print(ret)

    # 주문 버튼이 눌렸을 때의 로직함수입니다.
    def OnBtn4_clicked(self):
        print("주문직전")
        ret = self.kiwoom.SendOrder("주식주문", "0107", "3670956111", 1, "011690", 1, 1610, "00", "") # 유양디앤유
        #ret = self.kiwoom.SendOrder("주식주문", "0101", "3789954511", 1, "011690", 1, 1610, "0", "") # 유양디앤유
        #ret = self.kiwoom.SendOrder("주식주문", "0107", "8074249411", 1, "011690", 1, 1610, "0", "") # 유양디앤유
        
        #ret = self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",\
        #                ["주식주문", "0107", "3670956111", 1, "016090", 1, 3700, "0", ""])
        print("주문전송결과ret : {}".format(ret))

    # 검색조건 활성화 버튼
    def OnBtn5_clicked(self):
        ret = self.kiwoom.SetRealReg("0001", "032540", "10", "0") #TJ 미디어
        # 검색조건 등록 관련 클래스 선언 및 설정해주는 부분

        self.jongmok_set.add_jongmok("002620")  # 제일약품(신) 
        self.jongmok_set.add_jongmok("033320")  # 제이씨현시스템(신) 
        self.jongmok_set.add_jongmok("065450")  # 빅텍(신) 
        self.jongmok_set.add_jongmok("086060")  # 진바이오텍(코)

        # 종목별 logic을 추가합니다.
        self.jongmok_set.add_jongmok_logic("002620", Logic_Buy(self.jongmok_set.get_jongmok("002620"), 49700, 1))
        self.jongmok_set.add_jongmok_logic("033320", Logic_Buy(self.jongmok_set.get_jongmok("033320"), 5000, 1))
        self.jongmok_set.add_jongmok_logic("065450", Logic_Buy(self.jongmok_set.get_jongmok("065450"), 3150, 1))
        self.jongmok_set.add_jongmok_logic("086060", Logic_Buy(self.jongmok_set.get_jongmok("086060"), 4820, 1))

        # 키움 실시간 데이터 수신에 등록해줍니다.
        self.jongmok_set.register_realtime_all()


    # 분봉 차트 조회 버튼 클릭시 동작함수입니다
    def OnBtn6_clicked(self):
        self.jongmok_set.add_jongmok("002620")  # 제일약품(신) 

        self.kiwoom.SetInputValue("종목코드", "002620")
        self.kiwoom.SetInputValue("틱범위", "1")
        self.kiwoom.SetInputValue("수정주가구분", "0")
        self.kiwoom.CommRqData("1분봉차트요청", "opt10080", 0, "0101")

    # TR가격데이터가 수신되었을 때의 트리거 함수입니다.
    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext,\
            nDataLength, sErrorCode, sMessage, sSPlmMsg):
        if sRQName == "주식기본정보":
            print('come here')
            #cur_price = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", \
                                #sTRCode, "", sRQName, 0, "현재가")
            cur_price = self.kiwoom.CommGetData(sTRCode, "", sRQName, 0, "현재가").strip()
            cur_code = self.kiwoom.CommGetData(sTRCode, "", sRQName, 0, "종목코드").strip()
            cur_price = cur_price[1:]
            print("현재가 : {}, {}".format(sTRCode, cur_price))

            # 관심종목의 경우 조건만족 매수로직입니다.
            if cur_code == "032540" and self.ordered == 0: # TJ미디어, 아직 매수주문이 발생하지 않은 경우
                # 고가를 넘어섰을 때 30만원을 추가로 매수합니다.
                if int(cur_price) > 3765:
                    ret = self.kiwoom.SendOrder("주식주문", "0107", "3670956111", 1, "032540", 1, 3800, "00", "") # TJ미디어
                    self.ordered = 1

        elif sRQName == "1분봉차트요청":
            print('.1분봉요청정보 수신')
            code = self.kiwoom.CommGetData(sTRCode, "", sRQName, 0, "종목코드").strip()
            print('종목코드:{}'.format(code))

            arr = self.kiwoom.GetCommDataEx(sTRCode, sRQName)
            print("{}, {}, {}, ...".format(arr[0][0], arr[1][0], arr[2][0]))
            #self.jongmok_set.get_jongmok(code).prices[SP_1_chart] = arr
            self.jongmok_set.get_jongmok(code).set_price(SP_1_chart, arr)
            self.jongmok_set.get_jongmok(code).update_prices()

        else:
            pass
            #print('OnReceiveTrData::{},{},{},{},{},{},{},{}'.format(sRQName.strip(), sTRCode, sRecordName, sPreNext, \
            #                    nDataLength, sErrorCode, sMessage, sSPlmMsg))
            #order_vol = self.kiwoom.CommRqData("주문수량", "opt10012", 0, "0101")
            #print('주문수량 : {}'.format(order_vol))

    # 실시간 가격데이터가 수신되었을 때의 트리거 함수입니다.
    def OnReceiveRealData(self, code, sRealType, sRealData):
        print('실시간 시세 OnReceive!()::종목코드:{}'.format(code))
        #real_price = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", code, 0 )

        price = self.kiwoom.GetCommRealData("주식시세", 10).strip()[1:]
        try:
            if price != "":
                #self.jongmok_set.update_jongmok_price(code, int(price), "cur")
                self.jongmok_set.update_jongmok_price(code, int(price), VP_cur)

        except:
            print(".OnReceiveRealData::update_jongmok_price 중 에러발생")
        #real_price = self.kiwoom.GetCommRealData("주식시세", 10).strip()[1:]
        #real_volume = self.kiwoom.GetCommRealData("주식체결", 15).strip()
        #print("종목코드 : {}, 현재가 : {}, 거래량 : {}".format(code, real_price, real_volume))

        self.jongmok_set.unseal(code)

        '''
        # 관심종목의 경우 조건만족 매수로직입니다.
        if code == "032540" and self.ordered == 0: # TJ미디어, 아직 매수주문이 발생하지 않은 경우
            # 고가를 넘어섰을 때 30만원을 추가로 매수합니다.
            if int(real_price) > 3800:
                ret = self.kiwoom.SendOrder("주식주문", "0107", "3670956111", 1, "032540", 1, 3800, "00", "") # TJ미디어
                self.ordered = 1
        '''

    # 주문 접수/확인 수신 시 트리거되는 함수입니다.
    def OnReceiveChejanData(self, sGubun, nItemCnt, sFIDList):
        print("OnReceiveChejanData()")

        order_num = self.kiwoom.GetChejanData(302)
        order_vol = self.kiwoom.GetChejanData(900)
        order_price = self.kiwoom.GetChejanData(901)
        
        print("주문번호 : {}, 주문수량 : {}, 주문가격 : {}".format(order_num.strip(), order_vol, order_price))

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        print("OnReceiveMsg::sMsg:{}".format(sMsg))

    @asyncio.coroutine
    def timer_async(self):
        step = 1
        self.timer_dict = {}
        for i in range(60):
            self.timer_dict[i] = 0
        fetched = 0     # 매분 00 초에 명령수행이 완료되었는지 여부를 저장합니다
        tic_stack = 5   # 00초가 된 후 5틱(5초 이상) 동안은 실행명령을 한다. 다만 fetched 이면 패스가 되는 방식이다
        while 1:
            with tyUtils.time_elapsed(self.loop, self):
                print("1초 {}".format(self.loop.time()))

                sec = int(tyUtils.now().strftime("%M,%S")[3:])
                print(tyUtils.now().strftime("%M,%S"))

                # 00~05초 일 경우, (혹시나 피치못할 타이머 딜레이 발생을 감안하여 +- 5초를 감시대상으로 삼습니다)
                if sec < 5:
                    # tic stack 을 하나 빼주고 fetch를 실행합니다. 하지만 이미 fetched라면 패스가 됩니다.
                    if fetched == 0:
                        '''
                        process
                        '''
                        print('작업합니다^^ 헤헷')
                        fetched = 1
                # 55초 이상일 경우부터는 fetched 를 0으로 다시 초기화해줍니다
                elif sec >= 55:
                    fetched = 0

                #임시로 동작이 없어서 sleep 을 구겨넣어봤습니다.
                yield from asyncio.sleep(0.5)
                
                # some process
            #print('elapsed:{}'.format(self.elapsed))
            if self.elapsed < 1:
                step = 1 - self.elapsed
            else :
                step = 1
            #self.elapsed > 1 ? (step = 2 - self.elapsed) : step = 1
            #print('step:{}'.format(step))
            yield from asyncio.sleep(step)



with loop:
    window = MyWindow(loop)
    #asyncio.async(window.initialize())
    window.initialize()
    window.show()
    #loop.run_until_complete(window.initialize())
    #loop.call_soon(window.initialize())
    loop.run_forever()

'''
if __name__ == "__main__":
    QCoreApplication.setLibraryPaths([qt5_url])

    #app = QApplication(sys.argv)
    window = MyWindows()
    window.show()
    #app.exec_()
'''
