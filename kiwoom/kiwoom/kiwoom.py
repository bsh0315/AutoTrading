from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
class Kiwoom(QAxWidget):
    
    def __init__(self):
        super().__init__()
        print("키움 클래스 실행")
        
        #### eventloop 모음 ###
        self.login_event_loop = None
        ###############################
        
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConect()
        
    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOPENAPICtrl.1")
        
    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        
    def login_slot(self,errCode):
        print(errCode)
        
        self.login_event_loop.exit() # 키움 로그인 완료시에 이벤트 루프를 끊어줌
        
    def signal_login_commConect(self):
        self.dynamicCall("CommConnect()")
        
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()