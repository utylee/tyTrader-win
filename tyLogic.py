

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

    def update_jongmok_price(self, code, price):
        self.jongmok_dict[code]._setprice(price)

        #for i in self.jongmok_list:
        #    if i.code == code:
        #        i._setprice(price)

    def unseal(self, code):
        self.jongmok_dict[code].unseal()
        #for i in self.jongmok_list:
        #    if i.code == code:
        #        i.unseal()


class Jongmok:
    def __init__(self, ocx, code):
        self.ocx = ocx
        self.code = code
        self.price = 0
            
        self.logic_list = []

    def _add_logic(self, logic):
        self.logic_list.append(logic)

    def unseal(self):
        if len(self.logic_list) > 0:
            for i in self.logic_list:
                i.doLogic()

    def _setprice(self, price):
        self.price = price



class Logic:
    def __init__(self): pass
       #self.ocx = ocx 
    def doLogic(self) : pass


# 조건만족 실제 매수 구현클래스입니다
class Logic_Buy(Logic):
    def __init__(self, jongmok, buy_price, buy_volume):
        super().__init__()

        self.jongmok = jongmok 
        self.buy_price = buy_price              # 매수발동 가격 조건입니다.
        self.buy_volume = buy_volume                    # 매수 물량입니다.
        self.ordered = 0

    def doLogic(self):
        print(".매수감시::종목코드:{}, 설정가격:{}, 현재가:{}".format(self.jongmok.code, \
                            self.buy_price, self.jongmok.price)) 
        if self.jongmok.price > self.buy_price and self.ordered == 0:
            ret = self.jongmok.ocx.SendOrder("주식주문", "0107", "3670956111", 1, self.jongmok.code, self.buy_volume, \
                        self.buy_price + 100, "00", "")

            # 거래주문 성공시 중복주문 방지 플래그를 설정합니다
            if ret == 0:
                self.ordered = 1

