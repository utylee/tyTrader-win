

class Jongmok:
    def __init__(self, ocx, sCode):
        self.ocx = ocx
        self.sCode = sCode

    def accept(self, logic) : pass
        logic.doLogic()

class Logic:
    def __init__(self, ocx):
       self.ocx = ocx 
    def doLogic(self) : pass

class ConcreteStrategy:
    pass


class Logic_BuyCredit(Logic):
    def __init__(self):
        super().__init__()

    def doLogic(self):
        



class Buy_TJMedia(ConcreteStrategy)




