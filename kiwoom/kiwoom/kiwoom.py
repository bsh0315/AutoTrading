from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    
    def __init__(self):
        super().__init__()
        print("키움 클래스 실행")
        
        #### eventloop 모음 ###
        self.login_event_loop = None # 로그인 시
        self.detail_account_info_event_loop = None # 예수금 TR요청 시 사용
        self.detail_account_mystock_info_event_loop = None # 계좌평가잔고 요청시 사용
        ###############################
        
        
        ##### 변수 모음 ########
        self.account_num = None # 내 계좌번호
        self.screen_my_info = "2000"
        #######################
        
        #### 계좌관련 변수 #####
        self.use_money = 0 # 예수금
        self.use_money_percent = 0.7 # 예수금 중 실제로 매매에 사용할 비율
        ###########################
        
        
        
        
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConect()
        self.get_account_info() 
        self.detail_account_info() # 예수금 가져오는 곳
        self.detail_account_mystock() # 계좌평가 잔고 내역 요청
        
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
        
    def detail_account_mystock(self, sPrevNext = "0"):
        
        print("계좌평가잔고내역요청")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", 0, self.screen_my_info)
        
        # 요청 후 이벤트 루프 실행
        self.detail_account_mystock_info_event_loop = QEventLoop()
        self.detail_account_mystock_info_event_loop.exec_()
        
    
    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        print(f"TR 데이터 수신: sScrNo={sScrNo}, sRQName={sRQName}, sTrCode={sTrCode}, sPrevNext={sPrevNext}")
        
        '''
        TR요청을 받는 변수 
        sScrNo : 스크린번호
        sRQName : 내가 요청했을 때 지은 이름
        sTrCode : 요청id, tr코드
        sRecordName : 사용안함.
        sPrevNext : 다음 페이지가 있는지
        '''
        
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            print(f"예수금 : {int(deposit)}")
            
            self.use_money = int(deposit) * self.use_money_percent
            self.use_money /= 4

            
            avail_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print(f"출금가능금액 : {int(avail_money)}")
            
            avail_money1 = self.dynamicCall("GetCommData(String, String, int, string)", sTrCode, sRQName, 0, "d+1출금가능금액")
            avail_money2 = self.dynamicCall("GetCommData(String, String, int, string)", sTrCode, sRQName, 0, "d+2출금가능금액")
            print(f"d+1 출금가능금액 : {int(avail_money1)}\nd+2 출금가능금액 : {int(avail_money2)}")
            
            # 이벤트 루프 종료
            self.detail_account_info_event_loop.exit() 
            
        if sRQName == "계좌평가잔고내역요청":
            
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money = int(total_buy_money)
            print(f"총 매입금액 : {(total_buy_money)}")
            
            total_profit_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_rate = float(total_profit_rate)
            print(f"총 수익률 : {total_profit_rate}")
            '''
                #### 종목관련 변수 ######
                code = None # 종목코드
                stock_nm = None # 종목명
                stock_profit = None # 종목 평가손익
                stock_profit_rate = None # 종목 수익률 
                stock_quantity = None # 보유수량
                stock_avail_sell = None # 매매가능 수량
                stock_value = None # 종목 평가금액
                stock_position_rate = None # 전체 대비 보뮤비중
                ##################
            '''
            
            rows = self.dynamicCall("GetRepearCnt(QString, QString)", sTrCode, sRQName) # 보유종목 수
            # GetRepearCnt를 쓸때는 무조건 KOA studio의 TR에서 멀티데이터를 가져온다는 의미임.
            
            cnt = 0 
            print(f"보유종목 수 : {rows}")
            for i in range(rows):
                code  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, cnt, "종목번호")
                stock_nm  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, cnt, "종목명")
                stock_profit  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, cnt, "평가손익")
                stock_profit_rate  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, cnt, "수익률(%)")
                stock_quantity  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, cnt, "보유수량")
                stock_avail_sell  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, cnt, "매매가능수량")
                stock_value  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, cnt, "평가금액")
                stock_position_rate  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, cnt, "보유비중")
                
                
            
            self.detail_account_mystock_info_event_loop.exit()

            