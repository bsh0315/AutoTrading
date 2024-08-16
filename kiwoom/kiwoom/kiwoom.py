from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    
    def __init__(self):
        super().__init__()
        print("키움 클래스 실행")
        
        #### eventloop 모음 ###
        self.login_event_loop = None
        ###############################
        
        
        ##### 변수 모음 ########
        self.account_num = None # 내 계좌번호
        #######################
        
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConect()
        self.get_account_info()
        
    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOPENAPICtrl.1")
        
    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        
    def login_slot(self,errCode):
        print(errCode)
        errors(errCode)
        
        self.login_event_loop.exit() # 키움 로그인 완료시에 이벤트 루프를 끊어줌
        
    def signal_login_commConect(self):
        self.dynamicCall("CommConnect()")
        
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()
        
    def get_account_info(self):
        account_list = self.dynamicCall("GetLogininfo(String)", "ACCNO")
        
        self.account_num = account_list.split(';')[0] 
        #계좌가 여러개인 경우에 ;를 기준으로 나눠짐.
        # [0]은 제일 처음의 계좌번호를 가져옴
        
        print(f"나의 보유 계좌번호 : {self.account_num}")
