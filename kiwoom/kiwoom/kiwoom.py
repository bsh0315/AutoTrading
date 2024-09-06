from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import *
import os

class Kiwoom(QAxWidget):
    
    def __init__(self):
        super().__init__()
        print("키움 클래스 실행")
        
        #### eventloop 모음 ###
        self.login_event_loop = None # 로그인 시
        self.detail_account_info_event_loop = QEventLoop() #  TR(Transaction Request)요청 시 사용
        self.calcuator_event_loop = QEventLoop() # 종목 분석에 사용되는 이벤트 루프
        ###############################
        
        
        ##### 변수 모음 ########
        self.account_num = None # 내 계좌번호
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"
        self.portfolio_stock_dict = {}# 주식 포트폴리오 딕셔너리임.
        #######################
        
        #### 계좌관련 변수 #####
        self.use_money = 0 # 예수금
        self.use_money_percent = 0.7 # 예수금 중 실제로 매매에 사용할 비율
        self.buy_stock_quantity = 10 
        self.account_stock_dict = {} # 계좌 내 종목들의 딕셔너리
        self.not_account_stock_dict = {} # 계좌 내 미체결 주문에 대한 딕셔너리
        ###########################

        ##### 종목 분석용 변수 #####
        self.calcul_data = [] # 일봉 데이터가 들어가 있음
        ###########################
        
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConect()
        self.get_account_info() 
        self.detail_account_info() # 예수금 가져오는 곳
        self.detail_account_mystock() # 계좌평가 잔고 내역 요청
        self.not_concluded_account() # 실시간 미체결 요청
        
    
        #self.calculator_fnc() # 종목분석용 함수
        self.read_code()  # 저장된 종목들 불러온다.   
        
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
    
    def not_concluded_account(self, sPrevNext="0"): # 미체결 종목 정보 불러오기
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)
        
        # 요청 후 이벤트 루프 실행
        self.detail_account_info_event_loop.exec_()
        
    def day_kiwoom_db(self,code= None, date=None, sPrevNext = "0"):
        
        print("일봉데이터 요청")
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
        
        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", "date")
        
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식일봉차트조회", "opt10081", sPrevNext, self.screen_my_info)
        
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
                stock_position_rate  = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유비중(%)")
                
                code = code.strip()[1:]
                
                if code in self.account_stock_dict:
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
                asd = self.account_stock_dict[code]
                
                asd.update({"종목명" : stock_nm})
                asd.update({"평가손익" : stock_profit})
                asd.update({"수익률(%)" : stock_profit_rate})
                asd.update({"보유수량" : stock_quantity})
                asd.update({"매매가능수량" : stock_avail_sell})
                asd.update({"평가금액" : stock_value})
                asd.update({"보유비중(%)" : stock_position_rate})
                
                cnt += 1
            
            print(f"계좌에 가지고 있는 종목 수 : {cnt}")
            print(f"종목정보 : {self.account_stock_dict}")
            
            total_cnt += cnt
            
            if sPrevNext == 2:
                self.detail_account_mystock(sPrevNext = "2")
            else: 
                self.detail_account_info_event_loop.exit()
                
            print(f"총 종목 수 : {total_cnt}")
        
        elif sRQName == "실시간미체결요청":
            print("실시간 미체결 요청")
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
        
            if rows == 0:
                print("미체결 종목 없음")
                self.detail_account_info_event_loop.exit()
            else:
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

        elif "주식일봉차트조회" == sRQName:
            
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
            code = code.strip()
            print(f"일봉데이터 요청 {code}")
            
           

            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            print(f"받은 일봉 수 : {cnt}")
            
            # 한번 조회할 때마다 600일치를 조회가능함.
            for i in range(cnt): 
                data = []
                
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가") # 현재가는 종가 역할을 함.
                value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")
                trading_value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래대금")
                date = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "일자")
                starting_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")

                data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(starting_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                data.append("")
                
                self.calcul_data.append(data.copy()) 
                # copy() 함수는 객체의 얕은 복사를 만들 떄 사용함
                # 중첩된 리스트로써 둘 중 하나를 수정하면 다른 리스트도 수정된다.
            
            print(len(self.calcul_data))
            
            pass_success = False
            if sPrevNext == "2":
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
            else:
                print(f"총 일수 : {len(self.calcul_data)}")
                
                
                pass_success = False
                #120일 이평선을 그릴만큼의 데이터가 있는지 체크
                if self.calcul_data == None or len(self.calcul_data) < 120:
                    pass_success = False
                
                else : 
                    # 만약 120 이상의 일봉 데이터가 있다면
                    total_price = 0
                    for value in self.calcul_data[:120]: # 0부터 119까지. 즉 오늘부터 119일 전까지
                        total_price += int(value[1]) # 현재가(종가)를 모두 더함.
                        
                    average_line_120 = total_price/120
                    '''
                    calcul_data[][1] : 현재가
                    calcul_data[][2] : 거래량
                    calcul_data[][3] : 거래대금
                    calcul_data[][4] : 일자
                    calcul_data[][5] : 시가
                    calcul_data[][6] : 고가
                    calcul_data[][7] : 저가
                    '''
                    # 간단한 프로그램 매매 알고리즘, 그래빌의 매수신호 4법칙
                    # 오늘자 주가가 120 이평선에 걸쳐있는지를 확인함.
                    bottom_stock_price = False
                    check_price = None
                    if int(self.calcul_data[0][7]) <= average_line_120 and average_line_120 <= int(self.calcul_data[0][6]):
                        print("오늘 주가 120이평선에 걸쳐있는 것 확인")
                        bottom_stock_price = True # 120일 이평선에 걸쳐있는것 확인
                        check_price = int(self.calcul_data[0][6]) # 고가를 저장함.
                        
                        
                    # 과거 일봉들이 120일 이평선보다 밑에 있는지 확인.
                    # 만약 주가가 120일 이평선보다 위로 가게되면 계산이 진행됨.
                    prev_price = None # 과거의 일봉 저가를 의미함.
                    price_top_moving = False
                    if bottom_stock_price == True:
                        average_line_120 = 0
                        price_top_moving = False
                    
                    idx = 1
                    while(True):
                        if len(self.calcul_data[idx:1]) < 120: # 120일치의 일봉이 있는지 계속 확인함.
                            print("120일치가 없음")
                            break
                        total_price = 0
                        for value in self.calcul_data[idx:120+idx]:
                            total_price += int(value[1])
                            
                        prev_avg_line_120 = total_price/120
                
                        if prev_avg_line_120 <= int(self.calcul_data[idx][6]) and idx <= 20:
                            # 만약 20일 전에 고가가 120일 이평선과 같거나 위에 있으면 조건 통과 못함.
                            print("만약 20일 미만동안 고가가 120일 이평선과 같거나 위에 있으면 조건 통과 못함.")
                            price_top_moving =False
                            break
                        
                        elif int(self.calcul_data[idx][7] > prev_avg_line_120 and idx >20):
                            print("120일 이평선 위에 있는 일봉 확인됨")
                            price_top_moving = True
                            prev_price = int(self.calcul_data[idx][7])
                            break
                        idx +=1
                        
                        
                    # 해당 부분 이평선이 가장 최근 일자의 이평선 가격보다 낮은지 확인
                    if price_top_moving == True:
                        if average_line_120 > prev_avg_line_120:
                            print("포착된 이평선의 가격이 오늘자(최근일자) 이평선 가격보다 낮은 것 확인됨.")
                            print("포착된 부분의 일봉 저가가 오늘자 일봉의 고가보다 낮음.")
                            pass_success = True     
                            
                                               
            if pass_success == True:
                print("그랜빌 제 4법칙 매수신호 통과됨.")
                
                code_nm = self.dynamicCall("GetMasterCodeName(Qstring)", code) # 종목코드에 해당하는 종목명을 가져옴.
                      
                f = open("files/condition_stock.txt","a", encoding="utf8")  # "a"는 이어서 쓰기, "w"는 덮어쓰기    
                f.write(f"{code}\t{code_nm}\t{self.calcul_data[0][1]}\n")       
                f.close()
                
            elif pass_success == False:
                print("매수신호 통과 못함.")
                
                
                
                self.calcul_data.clear() # 사용 후 리스트의 내용물 지우기
                self.calcuator_event_loop.exit() # else: 문에서 원하는 조건에 해당하는 종목들을 미리 뽑아둠.
                
                
                
    def day_kiwoom_db(self,code= None, date=None, sPrevNext = "0"):
        QTest.qWait(3600) # 3.6초 딜레이
        
        print("일봉데이터 요청")
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
        
        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)
        # 만약 date = None 이라면 가장 최신 일자의 데이터부터 조회됨.
        
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식일봉차트조회", "opt10081", sPrevNext, self.screen_my_info)
        
        # 요청 후 이벤트 루프 실행
        self.calcuator_event_loop.exec()

            
    def get_code_list_market(self, market_code): # 특정 거래소의 종목코드들을 긁어옴.
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]
        
        return code_list
    
    def calculator_fnc(self): # 종목분석용 함수
        code_list = self.get_code_list_market("10") # 10은 코스닥 0은 코스피
        print(f"코스닥 종목 수 : {len(code_list)}")
        
        for idx, code in enumerate(code_list): # enumerate : 인덱스와 데이터를 같이 받음. enum(열거형)
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calculation_stock) # 스크린번호를 먼저 끊어줌, 기존의 데이터를 끊어줌.
            
            print(f"{idx+1} / {len(code_list)} : KOSDAQ Stock Code : {code} is updating... ")
            
            self.day_kiwoom_db(code=code)
            
            
    def read_code(self): # 저장된 종목들 불러오기
        if os.path.exists("files/condition_stock.txt"):
            f = open("files/condition_stock.txt", "r" ,encoding="utf8") # 텍스트 파일에 저장된 종목 정보를 라인별로 모두 가져옴.
            
            
            lines = f.readlines()
            for line in lines:
                if line != "":
                    ls = line.split("\t") # 종목코드 종목명 현재가로 구분 시킴.
                    
                    stock_code = ls[0]
                    stock_name = ls[1]
                    stock_price = int(ls[2].split("\n")[0])
                    stock_price = abs(stock_price) # 현재가에 -가 붙어있는 경우가 있음. 절댓값을 씌워줌.
                    
                    self.portfolio_stock_dict.update({stock_code : {"종목명" : stock_name}, "현재가" : stock_price})
                
                f.close()
                print(self.portfolio_stock_dict)