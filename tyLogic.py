
import asyncio

class Jongmok_set:
    def __init__(self, ocx):
        self.ocx = ocx
        self.jongmok_dict = {}

    def add_jongmok(self, code):
        jongmok = Jongmok(self.ocx, code) 
        #self.jongmok_dict.append(jongmok)
        self.jongmok_dict[code] = jongmok

    def add_jongmok_logic(self, code, logic):
        self.jongmok_dict[code]._add_logic(logic)

    def get_jongmok(self, code):
        return self.jongmok_dict[code]

    # 현재 소속된 종목들에 대해 실시간데이터 수신을 신청합니다.
    def register_realtime_all(self):
        # 종목코드들을 통해 신청문자열을 만들어냅니다.
        stringset = ""
        for i in self.jongmok_dict.items():
            stringset = "{}{};".format(stringset, i[1].code)
        print(".stringset : {}".format(stringset))
        ret = self.ocx.SetRealReg("0001", stringset, "10", "0") #금강철강
        print(".소속 종목들에 대해 실시간 데이터 수신을 신청했습니다:{}".format(ret))

    def update_jongmok_price(self, code, price, price_type):
        self.jongmok_dict[code]._setprice(price, price_type)

    def unseal(self, code):
        self.jongmok_dict[code].unseal()

#------------------------------------------------------------------------
class Jongmok:
    def __init__(self, ocx, code):
        self.ocx = ocx
        self.code = code
        #self.prices = {"cur": 0, "1_300": 0, "pivot_2": 0, "15_20": 0, "15_60":0, "15_120":0}
        self.prices = {}
        self._instance_connections = {}
        self.price = 0
        self._need_updated = []
            
        self.logic_list = []

        self.initialize()

    #def _add_price_type(self, obj_type, obj):
    def _add_price_type(self, obj_type):
        obj = obj_type()
        self._instance_connections[obj_type] = obj
        self.prices[obj_type] = obj.price

    def get_price(self, obj_type):
        return self._instance_connections[obj_type].price

    def set_price(self, obj_type, price):
        self._instance_connections[obj_type].price = price

    def initialize(self):
        self._add_price_type(VP_cur)
        self._add_price_type(SP_1_chart)
        self._add_price_type(VP_1_300)

    def _add_logic(self, logic):
        self.logic_list.append(logic)

    def unseal(self):
        if len(self.logic_list) > 0:
            for i in self.logic_list:
                i.doLogic()

    # 실시간 가격에 따라 변화해야하는 가격들을 업데이트해줍니다.
    def update_prices(self):
        for i in self.prices.keys():
            self._instance_connections[i].update_price(self)
            #self.prices[i].update_price(self)


    def _setprice(self, price, price_type):
        self.prices[price_type] = price


#------------------------------------------------------------------------
#------------------------------------------------------------------------
#------------------------------------------------------------------------

class Price:
    def __init__(self): 
        self.price = 0
        self.price_obj = None
    def fetch(self, jongmok): pass
    def update_price(self, jongmok): pass

#------------------------------------------------------------------------
class Variable_price(Price):
    def __init__(self):
        super().__init__()

    def update_price(self, jongmok):
        pass

class Static_price(Price):
    def __init__(self):
        super().__init__()

    #def fetch(self, jongmok): pass

#------------------------------------------------------------------------

class VP_1_300(Variable_price):
    def __init__(self):
        super().__init__()

    def update_price(self, jongmok):
        arr = jongmok.get_price(SP_1_chart)
        sum_price = 0
        for i in range(300):
            sum_price = sum_price + int(arr[i][0][1:])
        # 현재가를 즉각 반영하는 알고리즘은 추후 적용하기로 합니다
        #sum_price = jongmok.get_price[VP_cur]
        #for i in range(1, 300):
            #sum_price = sum_price + arr[i][0][1:] 
        self.price = int(sum_price / 300)
        print('self.price = {}'.format(self.price))


class VP_cur(Variable_price):
    def __init__(self):
        super().__init__()


class SP_1_chart(Static_price):
    def __init__(self):
        super().__init__()

    #def fetch(self, jongmok):
    #    jongmok.ocx.




class Logic:
    def __init__(self): pass
    def doLogic(self): pass


# 조건만족 실제 매수 구현클래스입니다
class Logic_Buy(Logic):
    def __init__(self, jongmok, buy_price, buy_volume):
        super().__init__()

        self.jongmok = jongmok 
        self.buy_price = buy_price              # 매수발동 가격 조건입니다.
        self.buy_volume = buy_volume            # 매수 물량입니다.
        self.ordered = 0

    def doLogic(self):
        print(".매수감시::종목코드:{}, 설정가격:{}, 현재가:{}".format(self.jongmok.code, \
                            #self.buy_price, self.jongmok.prices["cur"])) 
                            self.buy_price, self.jongmok.prices[VP_cur])) 
        #if self.jongmok.prices["cur"] > self.buy_price and self.ordered == 0:
        if self.jongmok.prices[VP_cur] > self.buy_price and self.ordered == 0:
            ret = self.jongmok.ocx.SendOrder("주식주문", "0107", "3670956111", 1, self.jongmok.code, self.buy_volume, \
                        self.buy_price + 100, "00", "")

            # 거래주문 성공시 중복주문 방지 플래그를 설정합니다
            if ret == 0:
                self.ordered = 1

