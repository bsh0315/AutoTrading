from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    
    def __init__(self):
        super().__init__()
        print("키움 클래스 실행")
        
        #### eventloop 모음 ###
        self.login_event_loop = None # 로그인 시
        self.detail_account_info_event_loop = QEventLoop() #  TR(Transaction Request)요청 시 사용
        ###############################
        
        
        ##### 변수 모음 ########
        self.account_num = None # 내 계좌번호
        self.screen_my_info = "2000"
        #######################
        
        #### 계좌관련 변수 #####
        self.use_money = 0 # 예수금
        self.use_money_percent = 0.7 # 예수금 중 실제로 매매에 사용할 비율
        self.buy_stock_quantity = 10 
        self.accout_stock_dict = None # 계좌 내 종목들의 딕셔너리
        self.not_account_stock_dict = None # 계좌 내 미체결 주문에 대한 딕셔너리
        ###########################

        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConect()
        self.get_account_info() 
        self.detail_account_info() # 예수금 가져오는 곳
        self.detail_account_mystock() # 계좌평가 잔고 내역 요청
        self.not_concluded_account() # 실시간 미체결 요청
        
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
        self.detail_account_info_event_loop.exec_()
    
    def not_concluded_account(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)
        
        # 요청 후 이벤트 루프 실행
        self.detail_account_info_event_loop.exec_()
        
        
    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        print(f"TR 데이터 수신: sScrNo={sScrNo}, sRQName={sRQName}, sTrCode={sTrCode}, sPrevNext={sPrevNext}")
        
        '''
        TR요청을 받는 변수 
        sScrNo : 스크린번호
        sRQName : 내가  TR을 요청할 때 부여한 이름
        sTrCode : 요청id, tr코드
        sRecordName : 사용안함.
        sPrevNext : 계좌 잔고에서 페이지당 최대 20개만 표기됨. 이를 넘어갈 경우 이 변수가 사용됨.
        '''
        
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            print(f"예수금 : {int(deposit)}")
            
            self.use_money = int(deposit) * self.use_money_percent
            self.use_money /= self.buy_stock_quantity # 한 종목을 매수할 때 이용가능한 시드의 1/self.buy_stock_quantity만큼 매수가능

            
            avail_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print(f"출금가능금액 : {int(avail_money)}")
            
            avail_money1 = self.dynamicCall("GetCommData(String, String, int, string)", sTrCode, sRQName, 0, "d+1출금가능금액")
            avail_money2 = self.dynamicCall("GetCommData(String, String, int, string)", sTrCode, sRQName, 0, "d+2출금가능금액")
            print(f"d+1 출금가능금액 : {int(avail_money1)}\nd+2 출금가능금액 : {int(avail_money2)}")
            
            # 이벤트 루프 종료
            self.detail_account_info_event_loop.exit() 
        
        total_cnt = 0
        if sRQName == "계좌평가잔고내역요청":
            
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money = int(total_buy_money)
            print(f"총 매입금액 : {(total_buy_money)}")
            
            total_profit_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_rate = float(total_profit_rate)
            print(f"총 수익률 : {total_profit_rate}")
            
            
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName) # 보유종목 수
            # GetRepearCnt를 쓸때는 무조건 KOA studio의 TR에서 멀티데이터를 가져온다는 의미임.
            # GetRepeatCnt 함수는 한번에 20개의 종목의 데이터만 불러올 수 있음.
            
            cnt = 0
            for i in range(rows):
                code  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                stock_nm  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_profit  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "평가손익")
                stock_profit_rate  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                stock_quantity  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                stock_avail_sell  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")
                stock_value  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "평가금액")
                stock_position_rate  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유비중")
                
                code = code.strip()[1:]
                
                if code in self.accout_stock_dict:
                    pass
                else:
                    self.account_stock_dict.update({code : {}})
                    
                stock_nm = stock_nm.strip()
                stock_profit = int(stock_profit.strip())
                stock_profit_rate = float(stock_profit_rate.strip())
                stock_quantity = int(stock_quantity.strip())
                stock_avail_sell = int(stock_avail_sell.strip())
                stock_value = int(stock_value.strip())
                stock_position_rate = float(stock_position_rate.strip())
                
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
                
                self.accout_stock_dict[code].update({"종목명" : stock_nm})
                self.accout_stock_dict[code].update({"평가손익" : stock_profit})
                self.accout_stock_dict[code].update({"수익률" : stock_profit_rate})
                self.accout_stock_dict[code].update({"보유수량" : stock_quantity})
                self.accout_stock_dict[code].update({"매매가능수량" : stock_avail_sell})
                self.accout_stock_dict[code].update({"평가금액" : stock_value})
                self.accout_stock_dict[code].update({"보유비중" : stock_position_rate})
                
                cnt += 1
            
            print(f"계좌에 가지고 있는 종목 수 : {cnt}")
            print(f"종목정보 : {self.accout_stock_dict}")
            
            total_cnt += cnt
            
            if sPrevNext == 2:
                self.detail_account_mystock(sPrevNext = "2")
            else: 
                self.detail_account_info_event_loop.exit()
                
            print(f"총 종목 수 : {total_cnt}")
        
        elif sRQName == "실시간미체결요청":
            print("실시간 미체결 요청")
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            
            for i in range(rows):
                code  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")
                stock_nm  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                order_no  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
                order_status  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태") # 접수, 확인, 체결
                order_quantity  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문수량")
                order_price  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
                order_gubun  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문구분") # 매도, 매수, 
                not_quantity  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
                ok_quantity  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결량")
                
                code = code.strip()[1:]
                stock_nm = stock_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-') # "-매도", "+매수" 형태이므로
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}
                
                nasd = self.not_account_stock_dict[order_no]
                
                nasd.update({"종목코드" : code})
                nasd.update({"종목명" : stock_nm})
                nasd.update({"주문상태" : order_status})
                nasd.update({"주문수량" : order_quantity})
                nasd.update({"주문가격" : order_price})
                nasd.update({"주문구분" : order_gubun})
                nasd.update({"미체결수량" : not_quantity})
                nasd.update({"체결량" : ok_quantity})
                
                print(f"미체결 종목 정보 : {self.not_account_stock_dict[order_no]}")
                
                self.detail_account_info_event_loop.exit()
                