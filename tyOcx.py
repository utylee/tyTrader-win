import sys
import asyncio
import time
import socket
import win32com
import MySQLdb as mdb

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

# 새로운 Windows 7 로 이주하면서 qt버전 업데이트 때문인지 폴더 마지막이 살짝 바뀐 것 같습니다
#qt5_url = "C:/cygwin/home/utylee/.virtualenvs/tyTrader-win/Lib/site-packages/PyQt5/plugins" 
qt5_url = "C:/Users/utylee/.virtualenvs/tyTrader-win/Lib/site-packages/PyQt5/Qt/plugins" 

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
        #self.setGeometry(1800,300,600,700)
        self.setGeometry(100,100,300,400)

        self.ordered = 0   # 주문 발생 여부, 중복 주문을 막기 위한 변수입니다
        self.start_timer = 0    # 타이머 시작명령(실은 미리 시작했습니다. 다시 말해 적용을 알리는) 플래그

        self.dbcon = 0 # db 연결후의 con을 저장해 놓습니다
        self.dbcur = 0 # db 연결후의 실행커서를 저장해 놓습니다

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

        #git-pricewin
        # 실시간 호가를 위한  추가 버튼을 생성합니다.
        self.btn3 = QPushButton("호가", self)
        self.btn3.move(140, 100)
        self.btn3.clicked.connect(self.OnBtnHoga_clicked)

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
        self.btn6 = QPushButton("minute candle", self)
        self.btn6.setGeometry(20, 250, 180, 50)
        self.btn6.clicked.connect(self.OnBtn6_clicked)

        # 키움OpenApi 접속 창을 띄웁니다.
        #self.kiwoom.dynamicCall("CommConnect()")
        print(".접속중입니다. 기다려주십시오...")

        #asyncio.async(self.connect_and_timer())
        #self.loop.run_until_complete(self.connect_and_timer())
        #self.loop.call_soon(self.connect_and_timer())
        #self.connecting()

        # 타이머를 시작합니다.
        #yield from asyncio.async(self.timer_async())

    @asyncio.coroutine
    def connect_and_timer(self):
        ret = self.kiwoom.CommConnect()
        print('counting 6 seconds..')
        yield from asyncio.sleep(6)
        print('starting asyncio timer...')
        asyncio.async(self.timer_async())
        asyncio.async(self.OnBtnHoga_clicked())
        asyncio.async(self.connect_db())

    @asyncio.coroutine
    def connect_db(self):
        try:
            # mysqldb(mariaDB)에 접속합니다
            # utf8 을 지정해 주는 게 포인트 입니다
            self.dbcon = mdb.connect('localhost', 'root', 'sksmsqnwk11', 'kiwoom', charset = 'utf8') 
            # 커서를 얻어옵니다
            self.dbcur = self.dbcon.cursor()
            print('DB연결성공')
        except:
            print('db exception occurred!!')
            pass

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

    # real time 호가 버튼이 눌렸을 때의 로직함수입니다
    @asyncio.coroutine
    def OnBtnHoga_clicked(self):
        #ret = self.kiwoom.dynamicCall("SetRealReg(QString, QString, int, QString)", "9999", "053260", 0, "0")
        #ret = self.kiwoom.dynamicCall("SetRealReg(QString, QString, int, QString)", "9999", "016090", 0, "0")
        #ret = self.kiwoom.SetRealReg("0001", "032540", "10", "0") #TJ 미디어

        #self.kiwoom.SetInputValue("종목코드", "041020")  
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", "000890")
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "주식기본정보",\
                                    "OPT10001", 0, "0101")
        #self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간호가",\
                                    #"OPT10004", 0, "0101")
        #self.kiwoom.CommRqData("실시간호가", "OPT10004", 0, "0101")
        #ret = self.kiwoom.SetRealReg("0001", "041020", "10", "0") #모헨즈 헉 6920 인피? /screenno, tcode, fid_list, reg_type
        #ret = self.kiwoom.SetRealReg("0001", "041020", "21", "0") #호가 시간/screenno, tcode, fid_list, reg_type
        #ret = self.kiwoom.SetRealReg("0001", "041020", "41", "0") #매도1호가/?screenno, tcode, fid_list, reg_type
        #ret = self.kiwoom.SetRealReg("0001", "041020", "15", "0") #체결 거래량 /screenno, tcode, fid_list, reg_type
        ret = 0
        print("HogaClicked 리턴값: {}".format(ret))
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
            print('TrData!!')
            #cur_price = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", \
                                #sTRCode, "", sRQName, 0, "현재가")
            cur_price = self.kiwoom.CommGetData(sTRCode, "", sRQName, 0, "현재가").strip()
            cur_code = self.kiwoom.CommGetData(sTRCode, "", sRQName, 0, "종목코드").strip()
            cur_name = self.kiwoom.CommGetData(sTRCode, "", sRQName, 0, "종목명").strip()
            cur_price = cur_price[1:]
            print("데이터 : {}, {}".format(cur_name, cur_price))
            print("db 삽입중..")
            state = "INSERT INTO basicinfo (code, title) VALUES (\'{}\', \'{}\')".format(cur_code, cur_name)
            #state = "INSERT INTO basicinfo (code, title) VALUES (\'{}\', \'abcd\');".format(cur_code)
            print(state)
            print("리턴값:{}".format(self.dbcur.execute(state)))
            # commit을 해야 db에 반영이 됩니다
            self.dbcon.commit()

            '''
            INSERT INTO kiwoom.basicinfo
            (code, title, market, price, daydiff, beforeprice, `open`, high, low, limitup, limitdown)
            VALUES('', '', 0, '', '', '', '', '', '', '', '');
            '''


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

        '''
        INSERT INTO kiwoom.hoga
        (name, curtime, price, volume, sell_total, buy_total, \ 
        sell1, sell2, sell3, sell4, sell5, sell6, sell7, sell8, sell9, sell10, \ 
        buy1, buy2, buy3, buy4, buy5, buy6, buy7, buy8, buy9, buy10, \ 
        sell1vol, sell2vol, sell3vol, sell4vol, sell5vol, sell6vol, sell7vol, sell8vol, sell9vol, sell10vol,\ 
        buy1vol, buy2vol, buy3vol, buy4vol, buy5vol, buy6vol, buy7vol, buy8vol, buy9vol, buy10vol, \ 
        sell1diff, sell2diff, sell3diff, sell4diff, sell5diff, sell6diff, sell7diff, sell8diff, sell9diff, sell10diff,\ 
        buy1diff, buy2diff, buy3diff, buy4diff, buy5diff, buy6diff, buy7diff, buy8diff, buy9diff, buy10diff, \ 
        pricediff, dayhigh, dayopen, daylow, dayvol, daymoney, `day`, clienttime)
        VALUES('', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 0);
        '''

        print('\n** 실시간 시세 OnReceive!()::종목:{}, 타입: {}, 데이타: ->\n\t {}'.format(code, sRealType, sRealData))
        #real_price = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", code, 0 )

        name = ''
        curtime = ''
        price = ''
        volume = ''
        sell_total = ''
        buy_total = ''
        sell1 = sell2 = sell3 = sell4 = sell5 = sell6 = sell7 = sell8 = sell9 = sell10 = ''
        sell1vol = sell2vol = sell3vol = sell4vol = sell5vol = sell6vol = sell7vol = sell8vol = sell9vol = sell10vol = ''
        sell1diff = sell2diff = sell3diff = sell4diff = sell5diff = sell6diff = sell7diff = sell8diff = sell9diff = sell10diff = ''
        buy1 = buy2 = buy3 = buy4 = buy5 = buy6 = buy7 = buy8 = buy9 = buy10 = ''
        buy1vol = buy2vol = buy3vol = buy4vol = buy5vol = buy6vol = buy7vol = buy8vol = buy9vol = buy10vol = ''
        buy1diff = buy2diff = buy3diff = buy4diff = buy5diff = buy6diff = buy7diff = buy8diff = buy9diff = buy10diff = ''
        dayhigh = daylow = dayopen = ''
        pricediff = ''
        daymoney = ''
        dayvol = ''
        day = ''
        clienttime = 0

        now = tyUtils.now()
        
        day = now.strftime('%m%d') 
        clienttime = now.strftime('%H%M%S%f')

        if (sRealType == "주식체결"):
            print('1111111')
        elif (sRealType == "주식호가잔량"):
            print('2222222')
        elif (sRealType == "주식예상체결"):
            pass
        elif (sRealType == "주식시간외호가"):
            pass
        #price = self.kiwoom.GetCommRealData("주식시세", 10).strip()[0:]   #가격
        price = self.kiwoom.GetCommRealData("실시간호가", 10).strip()   #가격
        time = self.kiwoom.GetCommRealData("실시간호가", 21).strip()[0:] #호가시간
        sell1 = self.kiwoom.GetCommRealData("실시간호가", 41).strip()[0:] #매도호가1
        buy1 = self.kiwoom.GetCommRealData("실시간호가", 51).strip()[0:] #매수호가1
        vol = self.kiwoom.GetCommRealData("실시간호가", 15).strip()[0:] #체결 거래량
        #print("실시간 수신 : 현재시간 = {}, 서버시간 = {} \n 가격 = {}, 체결량 = {}, 매도1호가 = {}, 매수1호가 = {}".format(clienttime, time, price, vol, sell1, buy1))
        '''
        try:
            if price != "":
                #self.jongmok_set.update_jongmok_price(code, int(price), "cur")
                self.jongmok_set.update_jongmok_price(code, int(price), VP_cur)

        except:
            print(".OnReceiveRealData::update_jongmok_price 중 에러발생")
        '''
        #real_price = self.kiwoom.GetCommRealData("주식시세", 10).strip()[1:]
        #real_volume = self.kiwoom.GetCommRealData("주식체결", 15).strip()
        #print("종목코드 : {}, 현재가 : {}, 거래량 : {}".format(code, real_price, real_volume))


        # 조건 로직 개봉 로직입니다
        # 로직 일단 없이 테스
        #self.jongmok_set.unseal(code)

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
        fetched = 0     # 매분 00 초에 명령수행이 완료되었는지 여부를 저장합니다

        while 1:
            with tyUtils.time_elapsed(self.loop, self):
                #print("1초 {}".format(self.loop.time()))

                minsec = tyUtils.now().strftime("%M,%S")

                #sec = int(tyUtils.now().strftime("%M,%S")[3:])
                sec = int(minsec[3:])
                #print(minsec)

                # 00~05초 일 경우, (혹시나 피치못할 타이머 딜레이 발생을 감안하여 +- 5초를 감시대상으로 삼습니다)
                if sec < 5:
                    # tic stack 을 하나 빼주고 fetch를 실행합니다. 하지만 이미 fetched라면 패스가 됩니다.
                    if fetched == 0:
                        '''
                        process
                        '''
                        #print('(1분경과) 작업중..')
                        fetched = 1
                # 55초 이상일 경우부터는 fetched 를 0으로 다시 초기화해주면서 다음 분을 준비합니다
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

class Test():
    def __init__(self, loop):
        pass

    pass


with loop:
    window = MyWindow(loop)
    test = Test(loop)
    #asyncio.async(window.initialize())
    window.initialize()
    window.show()
    loop.run_until_complete(window.connect_and_timer())
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
