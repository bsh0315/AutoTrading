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
        self.screen_my_info = "2000"
        #######################
        
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConect()
        self.get_account_info() 
        self.detail_account_info() # 예수금 가져오는 곳
        
    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOPENAPICtrl.1") # 키움 OpenAPI를 사용하기 위한 설정
        
    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot) # 로그인 상태를 저장시킴. '0'이 나와야만 정상적으로 로그인된 것.
        self.OnReceiveTrData.connect(self.trdata_slot) # TR 데이터를 받음.
        
        
    def login_slot(self,errCode):
        print(errCode)
        errors(errCode)
        
        self.login_event_loop.exit() # 키움 로그인 완료시에 이벤트 루프를 끊어줌
        
    def signal_login_commConect(self):
        self.dynamicCall("CommConnect()")
        
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_() 
        # 이벤트 루프 실행. 로그인이 완료될 때까지 다음 코드의 실행을 미룸.
        
    def get_account_info(self):
        account_list = self.dynamicCall("GetLogininfo(String)", "ACCNO")
        # 계좌 번호를 가져옴. ACCNO(account number)는 계좌 번호를 반환함.
        '''
        이외의 옵션
        "USER_ID" : 사용자 ID를 반환합니다.
        "USER_NAME" : 사용자 이름을 반환합니다.
        '''
        
        self.account_num = account_list.split(';')[0] 
        #계좌가 여러개인 경우에 ;를 기준으로 나눠짐.
        # [0]은 제일 처음의 계좌번호를 가져옴
        
        print(f"나의 보유 계좌번호 : {self.account_num}")
    
    def detail_account_info(self):
        print("예수금 가져오기")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        
        # 요청 후 이벤트 루프 실행
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "예수금상세현황요청", "opw00001", 0, self.screen_my_info)
        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()
    
    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        print(f"TR 데이터 수신: sScrNo={sScrNo}, sRQName={sRQName}, sTrCode={sTrCode}, sPrevNext={sPrevNext}")
        '''
        TR요청을 받는 함수 
        sScrNo : 스크린번호
        sRQName : 내가 요청했을 때 지은 이름
        sTrCode : 요청id, tr코드
        sRecordName : 사용안함.
        sPrevNext : 다음 페이지가 있는지
        '''
        
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            print(f"예수금 : {deposit}")
            
            avail_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print(f"출금가능금액 : {avail_money}")
            
            avail_money1 = self.dynamicCall("GetCommData(String, String, int, string)", sTrCode, sRQName, 0, "d+1출금가능금액")
            avail_money2 = self.dynamicCall("GetCommData(String, String, int, string)", sTrCode, sRQName, 0, "d+2출금가능금액")
            print(f"d+1 출금가능금액 : {avail_money1}\nd+2 출금가능금액 : {avail_money2}")
            
            # 이벤트 루프 종료
            self.detail_account_info_event_loop.exit()