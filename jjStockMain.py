import time
import webbrowser
import numpy as np
import xlwt
import requests
import bs4
import warnings
import math
import commonFunctions.RqFunc as func
import webview
import wx
import wx.html2

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *
from datetime import datetime, timedelta
from tkinter import *
from urllib import parse

# 현재 시스템 날짜 (date)
sYear = datetime.today().strftime('%Y')
sMonth = datetime.today().strftime('%m')
sDay = datetime.today().strftime('%d')

# 연속데이터 조회하기전 sleep 시간 설정
TR_REQ_TIME_INTERVAL = 0.2

# 정적 사이즈
OBJECT_WIDTH_MAX_SIZE = 1780

# 로우데이터 배열.
rowdatas = []

# 수급데이터 주체별 인덱스 추출 딕셔너리
juche_dic = {2:4, 7:6, 15:5, 20:7, 25:8, 30:9, 35:10, 40:11, 45:12, 50:13, 55:14, 60:15, 65:16}

# 수급분석표용 주체별 인덱스 추출 딕셔너리
juche_analysis_dic = {3: 2, 4: 7, 5: 15, 6: 20, 7: 25, 8: 30, 9: 35, 10: 40, 11: 45, 12: 50, 13: 55, 14: 60, 15: 65}

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # ------------------------------------------------- init ----------------------------------------------------- #

        # Kiwoom Login
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.dynamicCall("CommConnect()")
        self.kiwoom.OnReceiveTrData.connect(self._receive_tr_data)      # TR 응답 처리
        self.kiwoom.OnReceiveRealData.connect(self._receive_real_data)     # 실시간 응답 처리
        self.kiwoom.OnEventConnect.connect(self.event_connect)

        self.setWindowTitle("JJstock Analysis")
        self.setGeometry(20, 50, 1800, 900)

        # 조회버튼
        btn1 = QPushButton("조회", self)
        btn1.setGeometry(244, 10, 60, 30)
        btn1.clicked.connect(self.btn1_clicked)

        # 종목코드명 검색 인풋.
        self.jongmokCode = QLineEdit(self)
        self.jongmokCode.setText('')
        self.jongmokCode.move(18, 10)
        self.jongmokCode.setAlignment(Qt.AlignHCenter)  # 텍스트 가운데 정렬
        self.jongmokCode.setStyleSheet("QLineEdit{border:1px solid #B21016}")
        self.jongmokCode.textEdited.connect(self._get_code_by_autocomplete)
        # ime 모드를 한글로 바꿔보장 ㅠㅠ.

        # 날짜 표시 인풋
        self.cal_label = QLabel(self)
        self.cal_label.setGeometry(122, 11, 100, 28)
        self.cal_label.setStyleSheet("QLabel{border:1px solid #000000; background-color:#393939}") # 레이블 스타일링
        # 날짜 레이블에는 기본적으로 하루전 데이터를 셋팅해놓는다.
        self.cal_label.setText((datetime.today() + timedelta(days=-1)).strftime("%Y-%m-%d"))
        self.cal_label.setText('2013-07-03')    # 테스트

        # 날짜선택 버튼에 이미지를 심어보자.
        self.calIcon = QIcon("getcal.png")
        self.cal_btn = QPushButton(self)
        self.cal_btn.setIcon(self.calIcon)

        self.cal_btn.setGeometry(213, 11, 29, 28)
        self.cal_btn.setStyleSheet("QPushButton{background-color:black}")
        self.cal_btn.clicked.connect(self.cal_btn_clicked)

        # 관심종목 / 잔고 탭
        self.accountAndFavorite = AccountAndFavorite(self)
        self.accountAndFavorite.setGeometry(5, 40, 620, 380)

        # 실시간 조회 데이터 set / remove 버튼
        self.real_set_btn = QPushButton(self)
        self.real_set_btn.setText('잔고 set')
        self.real_set_btn.setGeometry(440, 53, 80, 26)
        self.real_set_btn.clicked.connect(self.account_real_set)
        self.real_cls_btn = QPushButton(self)
        self.real_cls_btn.setText('잔고 end')
        self.real_cls_btn.setGeometry(520, 53, 80, 26)
        self.real_cls_btn.clicked.connect(self.account_real_cls)


        # 수급 그래픽 데이터 테이블
        self.sugupGUIHeaders = ['', '개인', '세력합', '외국인', '금융투자', '보험', '투신', '기타금융', '은행', '연기금'\
                                , '사모펀드', '국가', '기타법인', '내외국인', '손바뀜율']
        self.sugupGUIColTitles = ['주가선도', '보유비중', '분산추이', '평균단가', '이동평균']
        self.sugupGUItable = QTableWidget(0, self.sugupGUIHeaders.__len__(), self)
        self.sugupGUItable.setHorizontalHeaderLabels(self.sugupGUIHeaders)
        self.sugupGUItable.setRowCount(5)
        for cnt in range(len(self.sugupGUIHeaders)):
            if cnt == 0 : self.sugupGUItable.setColumnWidth(cnt, 94)
            else: self.sugupGUItable.setColumnWidth(cnt, 76)
        for cnt2 in range(5):
            self.sugupGUItable.setRowHeight(cnt2, 26)
            self.sugupGUItable.setItem(cnt2, 0, QTableWidgetItem(self.sugupGUIColTitles[cnt2]))
            self.sugupGUItable.item(cnt2, 0).setTextAlignment(Qt.AlignCenter)
        self.sugupGUItable.setGeometry(620, 40, 1160, 156)      # 위치 및 크기 설정
        self.sugupGUItable.verticalHeader().setVisible(False)   # 번호 감춤.
        # 최초 데이터 merge
        self.sugupGUItable.setSpan(0, 1, 5, 14)
        self.sugupGUItable.setItem(0, 1, QTableWidgetItem('조회된 데이터가 없습니다.'))
        self.sugupGUItable.item(0, 1).setTextAlignment(Qt.AlignCenter)
        self.sugupGUItable.setSelectionMode(QAbstractItemView.NoSelection)  # 셀렉트 안됨.
        self.sugupGUItable.setEditTriggers(QAbstractItemView.NoEditTriggers) # 수정안됨.

        # 차트 탭메뉴
        self.chartTabWid = ChartTabWid(self)
        self.chartTabWid.setGeometry(608, 184, 1186, 496)

        # 리포트 and 뉴스 탭메뉴
        self.newsDataTabWid = NewsDataTabWid(self)
        self.newsDataTabWid.setGeometry(5, 400, 620, 280)

        # 로우데이터/수급분석표 탭메뉴
        self.rowDataTabWid = RowDataTabWid(self)
        self.rowDataTabWid.setGeometry(5, 660, OBJECT_WIDTH_MAX_SIZE+10, 240)

        # 종목별투자자기관별요청 데이터조회 테이블 확장 버튼 (컨트롤러의 선언순서에 따라 z-order가 달라진다)
        self.exp_dt_btn = QPushButton(self)
        self.exp_dt_btn.setGeometry(600, 680, 600, 18)
        self.exp_dt_btn.setText('▲')
        self.exp_dt_btn.clicked.connect(self.exp_dt_btn_clicked)

        # 엑셀 저장
        self.save_to_xls_btn = QPushButton(self)
        self.save_to_xls_btn.setGeometry(1704, 676, 80, 24)
        self.save_to_xls_btn.setText('엑셀저장')
        self.save_to_xls_btn.clicked.connect(self.savefile)

        # 종목코드 검색 된 리스트
        self.listWidgetSearched = QListWidget(self)
        self.listWidgetSearched.move(10, 40)
        self.listWidgetSearched.setFixedWidth(330)
        self.listWidgetSearched.setFixedHeight(0)
        self.listWidgetSearched.itemDoubleClicked.connect(self._code_item_clicked)

        # 종목코드레이블
        self.jongcodelbl = QLabel(self)
        self.jongcodelbl.setText('011170')      #테스트코드
        self.jongcodelbl.setGeometry(320, 10, 300, 30)

        # 캘린더위젯
        self.cal = QCalendarWidget(self)
        self.cal.setGridVisible(True)
        self.cal.setSelectedDate(datetime.today() + timedelta(days=-1))
        self.cal.setGeometry(18, 40, 260, 250)
        self.cal.clicked[QDate].connect(self.showDate)
        self.cal.hide()

        # 최초진입 로딩중 화면 구현
        self.firstLoading = QLabel(self)
        self.firstLoading.setGeometry(0, 0, 1800, 900)
        self.firstLoading.setAlignment(Qt.AlignCenter)
        self.firstLoading.setStyleSheet("QLabel{background-color:rgba(0, 0, 0, 0.7)}")
        self.firstLoading.setText('로딩중...')

        # 최초 로딩완료후 종목검색 인풋에 포커스
        self.jongmokCode.setFocus()

        # 데이터로딩중 화면 구현
        self.rowDataLoading = QLabel(self)
        self.rowDataLoading.setGeometry(0, 0, 0, 0)     # 우선은 안보임
        self.rowDataLoading.setAlignment(Qt.AlignCenter)
        self.rowDataLoading.setStyleSheet("QLabel{background-color:rgba(0, 0, 0, 0.7)}")
        # ------------------------------------------------- 끝 ----------------------------------------------------- #

    # 실시간 잔고 요청 버튼클릭
    def account_real_set(self):
        print('실시간 잔고 요청을 등록합니다.')
        self._set_real_reg("001", "064350;045390", "10;930;931", "0")

    # 실시간 데이터 요청 메서드
    def _set_real_reg(self, screen_no, code, fid, str_real_type):
        """
        :param screen_no:
        :param code:
        :param fid:
        :param str_real_type: 0이면 첫 실시간 데이터 요청, 1이면 추가등록
        :return:
        """
        # 실시간 데이터를 요청하기전 커넥션이 맺어져 있는지 확인한다.
        if not func._get_connection_state(self.kiwoom):
            self.alert('키움서버에 연결되어 있지 않습니다. 프로그램을 재 실행해 주십시오.')
        else:
            print('연결이 되어 있습니다.')

        self.kiwoom.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen_no, code, fid, str_real_type)

    # 실시간 응답 처리
    def _receive_real_data(self, code, realtype, realdata):
        # 실시간 응답처리 개발중 (실시간성이라 장중에만 테스트가 가능하여 장중에만 개발가능함...)
        if realtype == "주식체결":
            print(self._get_comm_real_data("045390", 10))

    # 실시간요청 응답 데이터
    def _get_comm_real_data(self, realtype, fid):
        self.kiwoom.dynamicCall("GetCommRealData(QString, int)", realtype, fid)

    # 실시간 잔고 요청 취소클릭
    def account_real_cls(self):
        print('실시간 요청 취소합니다.')
        self._set_real_remove("ALL", "ALL")

    # ------------------ 키 이벤트 오버라이딩 -----------------
    def keyPressEvent(self, event):
        # print(event.key())

        if event.key() == Qt.Key_Down:
            # print('키다운 이벤트 발생')
            if self.jongmokCode.hasFocus():
                self.listWidgetSearched.setFocus()

        # 엔터키 (사무실에서는 Key_Enter가 안먹음)
        if event.key() == 16777220 or event.key() == Qt.Key_Enter:
            # print('엔터키이벤트 발생')
            if self.listWidgetSearched.hasFocus():
                self._code_item_clicked(self.listWidgetSearched.currentItem())

    def alert(self, text):
        QMessageBox.about(self, "알림", text)

    def _code_item_clicked(self, item):
        setcode = ''
        for keyv, valv in self.code_list_dic.items():
            if valv == item.text():
                setcode = keyv

        self.jongmokCode.setText(item.text())
        self.jongcodelbl.setText(setcode)
        self.listWidgetSearched.setFixedHeight(0)
        self.btn1_clicked()     # 선택이되면 바로 조회 실행함.

    def _get_code_by_autocomplete(self):
        self.listWidgetSearched.clear()
        viewCodeList = []
        for slist in self.code_list_dic.values():
            if self.jongmokCode.text() in slist:
                viewCodeList.append(slist)
        if len(viewCodeList) * 24 >= 240:
            self.listWidgetSearched.setFixedHeight(240)
        else:
            self.listWidgetSearched.setFixedHeight(len(viewCodeList) * 24)
        self.listWidgetSearched.addItems(viewCodeList)

    def event_connect(self, err_code):
        if err_code == 0:       # 로그인 성공시 메소드
            print("로그인 성공")

            # 종목코드 리스트 조회 코스피
            coderet = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", ['0'])
            codelist = coderet.split(';')
            self.kospi_code_name_list = []

            self.code_list_dic = {}

            for x in codelist:
                name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [x])
                self.code_list_dic[x] = name

            # 종목코드 리스트 조회 코스닥
            coderet2 = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", ['10'])
            codelist2 = coderet2.split(';')
            for x2 in codelist2:
                name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [x2])
                self.code_list_dic[x2] = name

            print('종목코드 리스트 가져오기가 성공하였습니다.')
            self.firstLoading.setGeometry(0, 0, 0, 0)   # 로딩중화면 감춤.



    # 종목별투자자기관별 리스트 테이블 확장 버튼 클릭시.
    def exp_dt_btn_clicked(self):
        thei = self.rowDataTabWid.height()
        if thei == 240:
            self.rowDataTabWid.setGeometry(5, 400, OBJECT_WIDTH_MAX_SIZE+10, 500)
            self.exp_dt_btn.move(600, 420)
            self.exp_dt_btn.setText('▼')
            self.save_to_xls_btn.setGeometry(1704, 416, 80, 24)
        else:
            self.rowDataTabWid.setGeometry(5, 660, OBJECT_WIDTH_MAX_SIZE+10, 240)
            self.exp_dt_btn.move(600, 680)
            self.exp_dt_btn.setText('▲')
            self.save_to_xls_btn.setGeometry(1704, 676, 80, 24)

    # 날짜 레이블 클릭
    def cal_btn_clicked(self):
        if self.cal.isVisible():
            self.cal.hide()
        else:
            self.cal.show()

    # calendar를 클릭하면 선택한 날짜를 레이블에 표시함.
    def showDate(self, datein):
        print(datein.toString('yyyy-MM-dd'))
        self.cal_label.setText(datein.toString('yyyy-MM-dd'))
        if self.cal.isVisible():
            self.cal.hide()

    def btn1_clicked(self):

        if self.jongcodelbl.text() == '':
            QMessageBox.about(self, "JJStock", "종목을 검색해주세요")
            return

        # 조회전 기존 데이터 리셋.
        self.rowDataTabWid.dataTable.setRowCount(0)
        self.rowDataLoading.setGeometry(0, 0, 1800, 900)
        self.rowDataLoading.setText('로딩중 ')

        rowdatas.clear()
        print('data load started.', end='')
        # 종목별투자자기관별요청

        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "일자", self.cal_label.text().replace('-', ''))
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.jongcodelbl.text())
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "금액수량구분", "2")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "단위구분", "1")
        self._comm_rq_data("opt10059_req", "opt10059", 0, "0796")

        while self.remained_data:
            time.sleep(TR_REQ_TIME_INTERVAL)
            # 이전 tr에서 마지막으로 저장한 날짜를 셋팅함. (self.lasted_date)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "일자", self.lasted_date)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.jongcodelbl.text())
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "금액수량구분", "2")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "단위구분", "1")
            self._comm_rq_data("opt10059_req", "opt10059", 2, "0796")

    # TR 요청 메서드
    def _comm_rq_data(self, rqname, trcode, next, screen_no):

        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen_no)
        # 이벤트 루프 만들기
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    # TR 응답 처리
    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        # print('다음건수가 있습니까? ==> ',  next)

        if next == '2':     # 연속데이터
            self.remained_data = True
        else:
            self.remained_data = False

        if rqname == 'opt10059_req':
            self._opt10059_set(rqname, trcode)

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    # 실시간 데이터 중지 메서드
    def _set_real_remove(self, screen_no, code):
        # 통신상태 확인
        if not func._get_connection_state(self.kiwoom):
            self.alert('키움서버에 연결되어 있지 않습니다. 프로그램을 재 실행해 주십시오.')

        self.kiwoom.dynamicCall("SetRealRemove(QString, QString)", screen_no, code)

    # 종목별투자자별 리스트 응답 후 처리
    def _opt10059_set(self, rqname, trcode):
        data_cnt = self._get_repeat_cnt(trcode, rqname)

        for i in range(data_cnt):
            crrOfRow = self.rowDataTabWid.dataTable.rowCount()
            self.rowDataTabWid.dataTable.setRowCount(crrOfRow + 1)
            self.rowDataTabWid.dataTable.setRowHeight(crrOfRow, 10)
            colidx = 0
            one_row_arr = []
            for j in self.rowDataTabWid.jmTabColItemInfo:
                getdata = self._comm_get_data(trcode, "", rqname, i, j)
                one_row_arr.append(getdata)
                self.rowDataTabWid.dataTable.setItem(crrOfRow, colidx, QTableWidgetItem(getdata))
                if colidx not in [0, 17]:
                    self._set_cell_style(crrOfRow, colidx, self.rowDataTabWid.dataTable.item(crrOfRow, colidx).text(), self.rowDataTabWid.dataTable, 'N')
                colidx += 1
                self.lasted_date = self._comm_get_data(trcode, "", rqname, i, '일자')

            rowdatas.append(one_row_arr)

        if self.remained_data == False:
            print('data load ended.')
            self._make_sugup_data()
        else:
            if len(self.rowDataLoading.text()) < 12:
                self.rowDataLoading.setText(self.rowDataLoading.text() + '.')
            else:
                self.rowDataLoading.setText('로딩중 ')
            print('.', end='')

    # 데이터 가져오기 함수
    def _comm_get_data(self, code, real_type, field_name, index, item_name):
        ret = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", code, real_type, field_name, index, item_name)
        return ret.strip()

    # 멀티데이터 크기 가져오기
    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    # 테이블컬럼 스타일 설정
    def _set_cell_style(self, row, col, value, totab, formatyn):
        # TODO 스타일설정
        if formatyn == 'Y':
            if len(totab.item(row, col).text()) >= (4 if value[:1] == '+' or value[:1] != '-' else 5):
                totab.setItem(row, col, QTableWidgetItem(format(int(totab.item(row, col).text()), ',')))

        totab.item(row, col).setTextAlignment(Qt.AlignRight)

        if value != '0' and value != '':
            if value[:1] == '+' or value[:1] != '-':
                totab.item(row, col).setForeground(QColor(255, 127, 39))
            if value[:1] == '-':
                totab.item(row, col).setForeground(QColor(148, 216, 246))

    # 수급 데이터 만들기
    def _make_sugup_data(self):
        self.rowDataLoading.setText('수급데이터 만들기함수가 실행되었습니다.')
        print('수급데이터 만들기함수가 실행되었습니다.')
        # 로우데이터 데이터를 ndarray객체로 만든다. ndarray객체가 수치계산에 더 빠르게 때문이다.
        global np_row_data
        np_row_data = np.array(rowdatas)

        global np_sugup_data
        try:
            np_sugup_data = np.zeros((np_row_data.shape[0], 70), dtype=int)
        except Exception as ex:
            print('에러발생', ex)

        # 데이터역정렬 (역순계산 때문에)
        np_row_data = np.flipud(np_row_data)

        for i in range(np_row_data.shape[0]):
            for j in range(np_sugup_data.shape[1]):
                if j == 0: np_sugup_data[i, j] = np_row_data[i, j]         # 일자
                if j == 1: np_sugup_data[i, j] = np_row_data[i, j]         # 현재가
                ## 여기서부터는 함수로 가능
                ## 각 주체별 개별데이터 생성.
                if j in juche_dic.keys() : self._make_sugup_part_data(j, i)

        ## 세력합 5일, 20일, 60일 추세 generate
        np_sugup_data = np.flipud(np_sugup_data)
        n5idx = n20idx = n60idx = 0
        seryuk_arr = np_sugup_data[:, 9]

        for i in range(np_sugup_data.shape[0]):
            if i % 5 == 0:  n5idx = i
            if i % 20 == 0: n20idx = i
            if i % 60 == 0: n60idx = i

            np_sugup_data[i, 12] = np.mean(seryuk_arr[n5idx:n5idx+5])
            np_sugup_data[i, 13] = np.mean(seryuk_arr[n20idx:n20idx + 20])
            np_sugup_data[i, 14] = np.mean(seryuk_arr[n60idx:n60idx + 60])

            # 수급주체별데이터 출력
            # print(np_sugup_data[i])

        print('수급데이터가 정상적으로 생성 되었습니다.')

        # 데이터역정렬 (수급분석표 생성 때문에)
        np_row_data = np.flipud(np_row_data)
        self._make_sugup_analysis()

    # ----------------------------------------------- 수급분석표 만들기 ------------------------------------------------
    def _make_sugup_analysis(self):
        print('수급분석표 만들기함수가 실행되었습니다')
        self.rowDataLoading.setText('수급분석표 만들기함수가 실행되었습니다.')
        # 조회해온 데이터의 년도측정(역정렬 상태)
        start_year = '' + np_row_data[np_row_data[:, 1].size-1, 0]
        start_year = datetime.strptime(start_year, "%Y%m%d")
        start_year = start_year.strftime('%Y')

        # 행의 수 (년도별 합계) 기본행의 수는 18 + 년도별에 따라 추가됨
        rows_count = 18 + (int(sYear) - int(start_year))
        print('년도의 수 : ', int(sYear) - int(start_year), '/ 총 카운트 : ', rows_count)

        self.rowDataTabWid.sugupTable.setRowCount(rows_count)

        danga_arr = np.abs(np.array(np_row_data[:, 1], dtype=int))  # 현재가 (부호없는 정수)
        trade_arr = np.array(np_row_data[:, 17], dtype=int)         # 거래량

        # 수급분석표 세력합 데이터 generator
        seryukhap_arr = np.zeros(np_row_data.shape[0], dtype=int)
        global rowsum
        for ix in range(np_row_data.shape[0]):
            # print(np_row_data[ix])
            rowsum = 0
            for jx in [5, 7, 8, 9, 10, 11, 12, 13, 15]:
                rowsum += int(np_row_data[ix, jx])
            seryukhap_arr[ix] = rowsum
        # print('seryukhap_arr : ', seryukhap_arr)

        # 테이블 데이터 박기
        for i in range(rows_count):
            if i < 5:
                self.rowDataTabWid.sugupTable.setItem(i, 0, QTableWidgetItem(np_row_data[i, 0]))   # 일자
                self.rowDataTabWid.sugupTable.setItem(i, 1, QTableWidgetItem(np_row_data[i, 1]))   # 평균단가
                self.rowDataTabWid.sugupTable.setItem(i, 2, QTableWidgetItem(np_row_data[i, 17]))  # 거래량
                self.rowDataTabWid.sugupTable.setItem(i, 3, QTableWidgetItem(np_row_data[i, 4]))   # 개인
                self.rowDataTabWid.sugupTable.setItem(i, 4, QTableWidgetItem(str(seryukhap_arr[i])))    # 세력합
                self.rowDataTabWid.sugupTable.setItem(i, 5, QTableWidgetItem(np_row_data[i, 5]))   # 외국인

                # 외국인 이후~
                for sidx in range(6, 16):
                    self.rowDataTabWid.sugupTable.setItem(i, sidx, QTableWidgetItem(np_row_data[i, sidx+1]))

            if i in [5, 6, 7, 8]:
                self.rowDataTabWid.sugupTable.setItem(i, 0, QTableWidgetItem(str(i-4) + '주'))

                pd = np.mean(danga_arr[(((i - 4) * 5) - 5):((i - 4) * 5)])
                if math.isnan(pd): pd = 0
                self.rowDataTabWid.sugupTable.setItem(i, 1, QTableWidgetItem(str(np.int(pd))))  # 평균단가

                gr = np.sum(trade_arr[(((i - 4) * 5) - 5):((i - 4) * 5)])
                self.rowDataTabWid.sugupTable.setItem(i, 2, QTableWidgetItem(str(np.int(gr))))  # 거래량
                self.rowDataTabWid.sugupTable.setItem(i, 3, QTableWidgetItem(str(np.int(self._make_juche_array(4, 5, i, 4, int, 'sum')))))  # 개인
                self.rowDataTabWid.sugupTable.setItem(i, 4, QTableWidgetItem(str(np.sum(seryukhap_arr[(((i - 4) * 5) - 5):((i - 4) * 5)]))))  # 세력합
                self.rowDataTabWid.sugupTable.setItem(i, 5, QTableWidgetItem(str(np.int(self._make_juche_array(4, 5, i, 5, int, 'sum')))))  # 외국인

                # 외국인 이후 ~
                for sidx in range(6, 16):
                    self.rowDataTabWid.sugupTable.setItem(i, sidx, QTableWidgetItem(str(np.int(self._make_juche_array(4, 5, i, sidx+1, int, 'sum')))))

            if i in [9, 10, 11]:
                self.rowDataTabWid.sugupTable.setItem(i, 0, QTableWidgetItem(str(i - 8) + '달'))
                pd = np.mean(danga_arr[(((i - 8) * 20) - 20):((i - 8) * 20)])
                if math.isnan(pd): pd = 0
                self.rowDataTabWid.sugupTable.setItem(i, 1, QTableWidgetItem(str(np.int(pd))))  # 평균단가
                gr = np.sum(trade_arr[(((i - 8) * 20) - 20):((i - 8) * 20)])
                self.rowDataTabWid.sugupTable.setItem(i, 2, QTableWidgetItem(str(np.int(gr))))  # 거래량
                self.rowDataTabWid.sugupTable.setItem(i, 3, QTableWidgetItem(str(np.int(self._make_juche_array(8, 20, i, 4, int, 'sum')))))  # 개인
                self.rowDataTabWid.sugupTable.setItem(i, 4, QTableWidgetItem(str(np.sum(seryukhap_arr[(((i - 8) * 20) - 20):((i - 8) * 20)]))))  # 세력합
                self.rowDataTabWid.sugupTable.setItem(i, 5, QTableWidgetItem(str(np.int(self._make_juche_array(8, 20, i, 5, int, 'sum')))))  # 외국인

                # 외국인 이후 ~
                for sidx in range(6, 16):
                    self.rowDataTabWid.sugupTable.setItem(i, sidx, QTableWidgetItem(str(np.int(self._make_juche_array(8, 20, i, sidx + 1, int, 'sum')))))

            if i in [12, 13, 14, 15]:
                self.rowDataTabWid.sugupTable.setItem(i, 0, QTableWidgetItem(str(i - 11) + '분기'))

                pd = np.mean(danga_arr[(((i - 11) * 60) - 60):((i - 11) * 60)])
                if math.isnan(pd): pd = 0
                self.rowDataTabWid.sugupTable.setItem(i, 1, QTableWidgetItem(str(np.int(pd))))  # 평균단가
                gr = np.sum(trade_arr[(((i - 11) * 60) - 60):((i - 11) * 60)])
                self.rowDataTabWid.sugupTable.setItem(i, 2, QTableWidgetItem(str(np.int(gr))))  # 거래량
                self.rowDataTabWid.sugupTable.setItem(i, 3, QTableWidgetItem(str(np.int(self._make_juche_array(11, 60, i, 4, int, 'sum')))))  # 개인
                self.rowDataTabWid.sugupTable.setItem(i, 4, QTableWidgetItem(str(np.sum(seryukhap_arr[(((i - 11) * 60) - 60):((i - 11) * 60)]))))  # 세력합
                self.rowDataTabWid.sugupTable.setItem(i, 5, QTableWidgetItem(str(np.int(self._make_juche_array(11, 60, i, 5, int, 'sum')))))  # 외국인

                # 외국인 이후 ~
                for sidx in range(6, 16):
                    self.rowDataTabWid.sugupTable.setItem(i, sidx, QTableWidgetItem(str(np.int(self._make_juche_array(11, 60, i, sidx + 1, int, 'sum')))))

            if i >= 16 and i < rows_count-2:
                self.rowDataTabWid.sugupTable.setItem(i, 0, QTableWidgetItem(str(i - 15) + '년'))
                pd = np.mean(danga_arr[(((i - 15) * 240) - 240):((i - 15) * 240)])
                if math.isnan(pd): pd = 0
                self.rowDataTabWid.sugupTable.setItem(i, 1, QTableWidgetItem(str(np.int(pd))))  # 평균단가
                gr = np.sum(trade_arr[(((i - 15) *  240) - 240):((i - 15) * 240)])
                self.rowDataTabWid.sugupTable.setItem(i, 2, QTableWidgetItem(str(np.int(gr))))  # 거래량
                self.rowDataTabWid.sugupTable.setItem(i, 3, QTableWidgetItem(str(np.int(self._make_juche_array(15, 240, i, 4, int, 'sum')))))  # 개인
                self.rowDataTabWid.sugupTable.setItem(i, 4, QTableWidgetItem(str(np.sum(seryukhap_arr[(((i - 15) * 240) - 240):((i - 15) * 240)]))))  # 세력합
                self.rowDataTabWid.sugupTable.setItem(i, 5, QTableWidgetItem(str(np.int(self._make_juche_array(15, 240, i, 5, int, 'sum')))))  # 외국인

                # 외국인 이후 ~
                for sidx in range(6, 16):
                    self.rowDataTabWid.sugupTable.setItem(i, sidx, QTableWidgetItem(str(np.int(self._make_juche_array(15, 240, i, sidx + 1, int, 'sum')))))

            if i == rows_count-2:
                self.rowDataTabWid.sugupTable.setItem(i, 2, QTableWidgetItem("현재 보유량"))
            if i == rows_count-1:
                self.rowDataTabWid.sugupTable.setItem(i, 2, QTableWidgetItem("최대 보유량"))

            if i < rows_count-2:
                for cssidx in range(3, 16):
                    self._set_cell_style(i, cssidx, self.rowDataTabWid.sugupTable.item(i, cssidx).text(),self.rowDataTabWid.sugupTable, 'Y')

            self.rowDataTabWid.sugupTable.setRowHeight(i, 10)

        # 보유량 계산
        self._make_amount(rows_count)

        print('수급분석표 생성이 완료되었습니다. ')
        self.rowDataLoading.setText('수급분석표 만들기함수가 실행되었습니다.')
        self.rowDataLoading.setGeometry(0, 0, 0, 0)

        # 증권리포트 가져오기 함수 실행
        self.getReportWebCrawling()

        # 동시에 guiTable만들기 실행
        self._make_sugup_gui_datas()
    # -------------------------------------------- 수급분석표 만들기 END -----------------------------------------------

    # 다운로드 링크열기
    def getDownload(self, item):
        chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
        if item.column() == 4:
            webbrowser.get(chrome_path).open_new("http://hkconsensus.hankyung.com" + item.text())

    # 수급gui 테이블 만들기
    def _make_sugup_gui_datas(self):
        print('수급GUI 테이블 생성이 시작되었습니다.')
        self.sugupGUItable.clearSpans()

        total_boyu = 0
        for j in range(3, 16):
            if j not in [4]:
                total_boyu += int(self.rowDataTabWid.sugupTable.item(self.rowDataTabWid.sugupTable.rowCount() - 2, j).text())

        self.progress_arr = list()

        for i in range(13):
            self.progress = QProgressBar()
            self.progress.setFixedHeight(20)
            self.progress_arr.append(self.progress)

        self.progress_arr[0].setValue(round(int(self.rowDataTabWid.sugupTable.item(self.rowDataTabWid.sugupTable.rowCount() - 2,3).text()) / total_boyu * 100))
        self.sugupGUItable.setCellWidget(1, 1, self.progress_arr[0])   # 개인 보유비중 (보유량 / 전체보유량 * 100)   재활용가능?
        # 세력합은 보유비중 계산안함
        #self.progress_arr[1].setValue(round(int(self.rowDataTabWid.sugupTable.item(self.rowDataTabWid.sugupTable.rowCount() - 2,4).text()) / total_boyu * 100))
        #self.sugupGUItable.setCellWidget(1, 2, self.progress_arr[1])  # 세력합 보유비중 (현재보유랑 / 최대보유량 * 100)
        # 외국인 이후 보유비중 for
        for idx in range(2, 13):
            self.progress_arr[idx].setValue(round(int(self.rowDataTabWid.sugupTable.item(self.rowDataTabWid.sugupTable.rowCount() - 2, idx+3).text()) / total_boyu * 100))
            self.sugupGUItable.setCellWidget(1, idx+1, self.progress_arr[idx])  # 보유비중(현재보유랑 / 최대보유량 * 100)


    # 증권리포트 크롤링 함수
    def getReportWebCrawling(self):
        self.newsDataTabWid.reportDataTable.setRowCount(0)

        mainUrl = 'http://hkconsensus.hankyung.com/apps.analysis/analysis.list?'
        pSdate = datetime.now() - timedelta(days=360)       # 1년치 데이터만
        pSdate = pSdate.strftime('%Y-%m-%d')
        pEdate = sYear + '-' + sMonth + '-' + sDay
        paramsArr = []
        paramsArr.append('sdate=')
        paramsArr.append(pSdate)
        paramsArr.append('&edate=')
        paramsArr.append(pEdate)
        paramsArr.append('&now_page=1')
        paramsArr.append('&pagenum=1000')   # 가져오는 갯수. 거의 무한대로 지정해서 다 가져오자.
        paramsArr.append('&search_text=')
        paramsArr.append(parse.quote(str.encode(self.jongmokCode.text(), 'euc-kr')))
        url = mainUrl + ("".join(paramsArr))
        print("증권리포트 크롤링 url : ", url)

        # header 특히 User-Agent가 있어야 함
        report_url_headers = {'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
                              'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.84 Safari/537.36'
                              }

        source_code = requests.get(url, headers=report_url_headers)
        source_code.encoding = 'euc-kr'
        # print(source_code.text)
        bs = bs4.BeautifulSoup(source_code.text, 'html.parser')
        mbody = bs.select('div.table_style01 table tbody tr')

        self.newsDataTabWid.reportDataTable.itemDoubleClicked.connect(self.getDownload)


        for ix in range(len(mbody)):
            crrOfRow = self.newsDataTabWid.reportDataTable.rowCount()
            one_row = mbody[ix].find_all('td')
            # 테이블 한줄 생성.
            self.newsDataTabWid.reportDataTable.setRowCount(crrOfRow + 1)
            self.newsDataTabWid.reportDataTable.setRowHeight(crrOfRow, 10)
            self.newsDataTabWid.reportDataTable.setItem(crrOfRow, 0, QTableWidgetItem(one_row[0].getText()))    # 작성일
            self.newsDataTabWid.reportDataTable.setItem(crrOfRow, 1, QTableWidgetItem(one_row[1].getText()))    # 분류
            self.newsDataTabWid.reportDataTable.setItem(crrOfRow, 2, QTableWidgetItem(one_row[2].getText()))    # 제목
            self.newsDataTabWid.reportDataTable.setItem(crrOfRow, 3, QTableWidgetItem(one_row[4].getText()))    # 출처(3인덱스는 담당자명)
            self.newsDataTabWid.reportDataTable.setItem(crrOfRow, 4, QTableWidgetItem(one_row[5].find('a').get('href')))    # 첨부파일

        print("증권리포트 크롤링이 완료되었습니다.")

    # 수급테이블 2차원 배열로부터 특정 열데이터를 추출하여 배열생성
    def _make_juche_array(self, minuspoint, period, rowidx, createcolidx, dtypes, mathtool):
        returnarr = np.array(np_row_data[:, createcolidx], dtype=dtypes)
        if mathtool == 'mean':
            self.rtnnp = np.mean(returnarr[(((rowidx - minuspoint) * period) - period):((rowidx - minuspoint) * period)])
            if math.isnan(self.rtnnp): self.rtnnp = 0
            return self.rtnnp
        elif mathtool == 'sum':
            return np.sum(returnarr[(((rowidx - minuspoint) * period) - period):((rowidx - minuspoint) * period)])

    # 보유량 계산
    def _make_amount(self, rows_count):
        print('보유량 계산중....')
        for idx in range(3, 16):
            self.rowDataTabWid.sugupTable.setItem(rows_count - 2, idx, QTableWidgetItem(str(np_sugup_data[0, juche_analysis_dic[idx] + 2])))    # 현재 보유량
            self.rowDataTabWid.sugupTable.setItem(rows_count - 1, idx, QTableWidgetItem(str(np_sugup_data[0, juche_analysis_dic[idx] + 3])))    # 최대 보유량(매집고점)

    # 수급 주체별 데이터 generator
    def _make_sugup_part_data(self, fromidx, rowidx):
        for i in range(fromidx, (fromidx+6)):

            # 세력 순매수 물량의 총합.
            stock_data = int(np_row_data[rowidx, 5]) + int(np_row_data[rowidx, 7]) + int(np_row_data[rowidx, 8]) + \
                         int(np_row_data[rowidx, 9]) + int(np_row_data[rowidx, 10]) + int(np_row_data[rowidx, 11]) + \
                         int(np_row_data[rowidx, 12]) + int(np_row_data[rowidx, 13]) + int(np_row_data[rowidx, 14]) + \
                         int(np_row_data[rowidx, 15])

            # 누적합
            if i == fromidx:
                if rowidx == 0:
                    # 삼항연산자 a if test else b
                    np_sugup_data[rowidx, i] = stock_data if fromidx == 7 else int(np_row_data[rowidx, juche_dic[fromidx]])
                if rowidx > 0:
                    np_sugup_data[rowidx, i] = np_sugup_data[rowidx-1, i] + (stock_data if fromidx == 7 else int(np_row_data[rowidx, juche_dic[fromidx]]))
            # 최고저점
            if i == fromidx+1:
                if rowidx == 0:
                    np_sugup_data[rowidx, i]= np_sugup_data[rowidx, i-1]
                elif rowidx > 0:
                    np_sugup_data[rowidx, i] = min(np_sugup_data[rowidx-1, i], np_sugup_data[rowidx, i-1])
            # 매집수량
            if i == fromidx+2:
                np_sugup_data[rowidx, i] = np_sugup_data[rowidx, i-2] - np_sugup_data[rowidx, i-1]
            # 매집고점
            if i == fromidx+3:
                if rowidx == 0:
                    np_sugup_data[rowidx, i] = np_sugup_data[rowidx, i-1]
                elif rowidx > 0:
                    np_sugup_data[rowidx, i] = max(np_sugup_data[rowidx-1, i], np_sugup_data[rowidx, i-1])
            # 분산비율
            if i == fromidx+4:
                if np_sugup_data[rowidx, i-2] == 0 or np_sugup_data[rowidx, i-1] == 0:
                    np_sugup_data[rowidx, i] = 0
                else:
                    np_sugup_data[rowidx, i] = (np_sugup_data[rowidx, i-2] / np_sugup_data[rowidx, i-1]) * 100

    # 엑셀파일로 데이터 저장
    def savefile(self):
        if self.rowDataTabWid.dataTable.rowCount() == 0:
            self.alert('대상 데이터가 로드되지 않았습니다. 종목검색을 먼저 실행해주십시오.')
            return

        ## 수정합시다으아~
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet("sheet 1", cell_overwrite_ok=True)
        self.add2(sheet)
        wbk.save("/Users/pconn/Desktop/export.xls")

    def add2(self, sheet):
        for currentColumn in range(self.rowDataTabWid.dataTable.columnCount()):
            for currentRow in range(self.rowDataTabWid.dataTable.rowCount()):
                try:
                    teext = str(self.rowDataTabWid.dataTable.item(currentRow, currentColumn).text())
                    sheet.write(currentRow, currentColumn, teext)
                except AttributeError:
                    pass
        self.alert('엑셀데이터 생성이 완료되었습니다.')


# 관심종목/실시간잔고 탭
class AccountAndFavorite(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.layout = QVBoxLayout(self)

        # 탭 스크린 초기화
        self.accTabs = QTabWidget()
        self.accTab1 = QWidget()
        self.accTab2 = QWidget()

        # 탭 추가
        self.accTabs.addTab(self.accTab1, "관심종목")
        self.accTabs.addTab(self.accTab2, "잔고")

        # 잔고 탭 화면 구성
        self.acctableHeaders = ['종목명', '현재가', '매입가', '보유수량', '가능수량', '평가손익', '수익률']
        self.accListTable = QTableWidget(0, self.acctableHeaders.__len__(), self)
        self.accListTable.setHorizontalHeaderLabels(self.acctableHeaders)

        self.accTab2.layout = QVBoxLayout(self)
        self.accTab2.layout.addWidget(self.accListTable)
        self.accTab2.setLayout(self.accTab2.layout)

        self.layout.addWidget(self.accTabs)
        self.setLayout(self.layout)


# 뉴스/리포트 탭 메뉴 규성
class NewsDataTabWid(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.layout = QVBoxLayout(self)

        # 탭 스크린 초기화
        self.newsTabs = QTabWidget()
        self.newsTab1 = QWidget()
        self.newsTab2 = QWidget()

        # 탭 추가
        self.newsTabs.addTab(self.newsTab1, "증권리포트")
        self.newsTabs.addTab(self.newsTab2, "뉴스")

        # 증권리포트 탭 내용 생성
        self.reportHeaders = ['작성일', '분류', '제목', '출처', '첨부']
        self.reportDataTable = QTableWidget(0, self.reportHeaders.__len__(), self)
        # self.reportDataTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.reportDataTable.setHorizontalHeaderLabels(self.reportHeaders)
        self.reportDataTable.setColumnWidth(0, 90)
        self.reportDataTable.setColumnWidth(1, 46)
        self.reportDataTable.setColumnWidth(2, 280)
        self.reportDataTable.setColumnWidth(3, 100)
        self.reportDataTable.setColumnWidth(4, 40)
        self.reportDataTable.setRowCount(1)
        self.reportDataTable.setRowHeight(0, 10)
        self.reportDataTable.setItem(0, 0, QTableWidgetItem('조회된 데이터가 없습니다.'))
        self.reportDataTable.setSpan(0, 0, 1, 5)     # setSpan을 걸때 count들은 최소 1이상이다.
        self.reportDataTable.item(0, 0).setTextAlignment(Qt.AlignCenter)
        self.reportDataTable.verticalHeader().setVisible(False)  # 번호 감춤.

        self.newsTab1.layout = QVBoxLayout(self)
        self.newsTab1.layout.addWidget(self.reportDataTable)
        self.newsTab1.setLayout(self.newsTab1.layout)

        # 뉴스 탭 내용 생성
        self.newsDataTable = QTableWidget(0, 5, self)
        self.newsTab2.layout = QVBoxLayout(self)
        self.newsTab2.layout.addWidget(self.newsDataTable)
        self.newsTab2.setLayout(self.newsTab2.layout)

        self.layout.addWidget(self.newsTabs)
        self.setLayout(self.layout)

class ChartTabWid(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.layout = QVBoxLayout(self)

        # 탭 스크린 초기화
        self.chartTabs = QTabWidget()
        self.chartTab1 = QWidget()
        self.chartTab2 = QWidget()
        self.chartTab3 = QWidget()
        self.chartTab4 = QWidget()

        # 탭 추가
        self.chartTabs.addTab(self.chartTab1, "매집현황")
        self.chartTabs.addTab(self.chartTab2, "분산비율")
        self.chartTabs.addTab(self.chartTab3, "투자자추이")
        self.chartTabs.addTab(self.chartTab4, "네이버차트")

        # 네이버차트 웹뷰
        self.chartTab4.layout = QVBoxLayout(self)
        app = wx.App()
        dialog = MyBrowser(None, -1)
        dialog.Show()
        # self.chartTab4.layout.addWidget(self.naverwebview)
        # self.chartTab4.setLayout(self.chartTab4.layout)

        # 레이아웃 바인딩
        self.layout.addWidget(self.chartTabs)
        self.setLayout(self.layout)

class MyBrowser(wx.Dialog):
  def __init__(parent, *args, **kwds):
    wx.Dialog.__init__(parent, *args, **kwds)
    sizer = wx.BoxSizer(wx.VERTICAL)
    parent.browser = wx.html2.WebView.New(parent)
    sizer.Add(parent.browser, 1, wx.EXPAND, 10)
    parent.SetSizer(sizer)
    parent.SetSize((700, 700))

# PyQt5의 QTableWidget을 이용한 탭메뉴 구성
class RowDataTabWid(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.layout = QVBoxLayout(self)

        # 탭 스크린 초기화
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()

        # 탭 추가
        self.tabs.addTab(self.tab1, "로우데이터")
        self.tabs.addTab(self.tab2, "수급분석표")

        # 첫번째 탭 내용 생성
        self.tab1.layout = QVBoxLayout(self)

        # 종목별투자자기관별요청 데이터조회 테이블
        # 테이블 헤더설정
        self.headers = ['일자', '현재가', '전일대비', '등락율', '개인', '외국인', '기관계', '금융투자', '보험', '투신',
                        '기타금융', '은행', '연기금등', '사모펀드', '국가', '기타법인', '내외국인', '거래량']

        # 종목별투자자기관별요청 컬럼정보
        self.jmTabColItemInfo = ['일자', '현재가', '전일대비', '등락율', '개인투자자', '외국인투자자', '기관계', '금융투자',
                                 '보험', '투신', '기타금융', '은행', '연기금등', '사모펀드', '국가', '기타법인', '내외국인',
                                 '누적거래대금']

        self.dataTable = QTableWidget(0, self.headers.__len__(), self)
        self.dataTable.setRowHeight(0, 10)
        self.dataTable.setHorizontalHeaderLabels(self.headers)

        self.tab1.layout.addWidget(self.dataTable)
        self.tab1.setLayout(self.tab1.layout)

        # 두번째 탭 내용 생성 (데이터 가공 수급데이터)
        self.tab2.layout = QVBoxLayout(self)
        self.sugupHeaders = ['일자', '평균단가', '거래량', '개인', '세력합', '외국인', '금융투자', '보험', '투신',
                             '기타금융', '은행', '연기금', '사모펀드', '국가', '기타법인', '내외국인']

        self.sugupTable = QTableWidget(0, self.sugupHeaders.__len__(), self)
        self.sugupTable.setRowHeight(0, 10)
        self.sugupTable.setHorizontalHeaderLabels(self.sugupHeaders)

        self.tab2.layout.addWidget(self.sugupTable)
        self.tab2.setLayout(self.tab2.layout)

        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)

    @pyqtSlot()
    def on_click(self):
        print("\n")
        for currentQTableWidgetItem in self.tableWidget.selectedItems():
            print(currentQTableWidgetItem.row(), currentQTableWidgetItem.column(), currentQTableWidgetItem.text())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    # ------------------------- 스타일 테마설정 -----------------------------
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(53, 53, 53))
    palette.setColor(QPalette.WindowText, Qt.white)
    palette.setColor(QPalette.Base, QColor(15, 15, 15))
    palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ToolTipBase, Qt.white)
    palette.setColor(QPalette.ToolTipText, Qt.white)
    palette.setColor(QPalette.Text, Qt.white)
    palette.setColor(QPalette.Button, QColor(53, 53, 53))
    palette.setColor(QPalette.ButtonText, Qt.white)
    palette.setColor(QPalette.BrightText, Qt.red)

    palette.setColor(QPalette.Highlight, QColor(142, 45, 197).lighter())
    palette.setColor(QPalette.HighlightedText, Qt.black)
    app.setPalette(palette)
    # ------------------------- 스타일 테마설정 끝 -----------------------------

    warnings.simplefilter("ignore")
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
