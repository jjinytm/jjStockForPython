import sys
import time
import numpy as np
import xlwt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *
from datetime import datetime
from tkinter import *
from commonFunctions import *

# 현재 시스템 날짜 (date)
sYear = datetime.today().strftime('%Y')
sMonth = datetime.today().strftime('%m')
sDay = datetime.today().strftime('%d')

# 연속데이터 조회하기전 sleep 시간 설정
TR_REQ_TIME_INTERVAL = 0.04

# 정적 사이즈
OBJECT_WIDTH_MAX_SIZE = 1780

# 로우데이터 배열.
rowdatas = []


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Kiwoom Login
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.dynamicCall("CommConnect()")
        self.kiwoom.OnReceiveTrData.connect(self._receive_tr_data)

        self.setWindowTitle("JJstock")
        self.setGeometry(20, 50, 1800, 900)

        # 조회버튼
        btn1 = QPushButton("조회", self)
        btn1.setGeometry(244, 10, 60, 30)
        btn1.clicked.connect(self.btn1_clicked)

        # 종목코드 입력 인풋.
        self.jongmokCode = QLineEdit(self)
        self.jongmokCode.setText('089590')  ## 임시종목 후성.
        self.jongmokCode.move(10,10)
        self.jongmokCode.setAlignment(Qt.AlignHCenter)  # 텍스트 가운데 정렬

        # 날짜 표시 인풋
        self.cal_label = QLabel(self)
        self.cal_label.setGeometry(115, 11, 100, 28)
        self.cal_label.setStyleSheet("QLabel{border:1px solid blue; background-color:#FAFAFA}") # 레이블 스타일링
        # 날짜 레이블에는 기본적으로 오늘 데이터를 셋팅해놓는다.
        self.cal_label.setText(sYear+'-'+sMonth+'-'+sDay)

        # 날짜선택 버튼에 이미지를 심어보자.
        self.calIcon = QIcon("getcal.png")
        self.cal_btn = QPushButton(self)
        self.cal_btn.setIcon(self.calIcon)

        self.cal_btn.setGeometry(213, 10, 30, 30)
        self.cal_btn.setStyleSheet("QPushButton{background-color:black}")
        self.cal_btn.clicked.connect(self.cal_btn_clicked)

        self.cal = QCalendarWidget(self)
        self.cal.setGridVisible(True)
        self.cal.setGeometry(10, 40, 260, 250)
        self.cal.clicked[QDate].connect(self.showDate)
        self.cal.hide()

        # 탭메뉴 추가
        self.rowDataTabWid = RowDataTabWid(self)
        self.rowDataTabWid.setGeometry(5, 660, OBJECT_WIDTH_MAX_SIZE+10, 240)

        # 종목별투자자기관별요청 데이터조회 테이블 확장 버튼 (컨트롤러의 선언순서에 따라 z-order가 달라진다)
        self.exp_dt_btn = QPushButton(self)
        self.exp_dt_btn.setGeometry(600, 677, 600, 18)
        self.exp_dt_btn.setText('▲')
        self.exp_dt_btn.clicked.connect(self.exp_dt_btn_clicked)

        # 엑셀 저장
        self.save_to_xls_btn = QPushButton(self)
        self.save_to_xls_btn.setGeometry(1680, 677, 80, 24)
        self.save_to_xls_btn.setText('엑셀저장')
        self.save_to_xls_btn.clicked.connect(self.savefile)

    # 종목별투자자기관별 리스트 테이블 확장 버튼 클릭시.
    def exp_dt_btn_clicked(self):
        thei = self.rowDataTabWid.height()
        if thei == 240:
            self.rowDataTabWid.setGeometry(5, 500, OBJECT_WIDTH_MAX_SIZE+10, 400)
            self.exp_dt_btn.move(600, 517)
            self.exp_dt_btn.setText('▼')
            self.save_to_xls_btn.setGeometry(1680, 517, 80, 24)
        else:
            self.rowDataTabWid.setGeometry(5, 660, OBJECT_WIDTH_MAX_SIZE+10, 240)
            self.exp_dt_btn.move(600, 677)
            self.exp_dt_btn.setText('▲')
            self.save_to_xls_btn.setGeometry(1680, 677, 80, 24)

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
        # 조회전 기존 데이터 리셋.
        self.rowDataTabWid.dataTable.setRowCount(0)
        rowdatas.clear()
        print('data load started.', end='')
        # 종목별투자자기관별요청
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "일자", self.cal_label.text().replace('-', ''))
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.jongmokCode.text())
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "금액수량구분", "2")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "단위구분", "1")
        self._comm_rq_data("opt10059_req", "opt10059", 0, "0796")

        while self.remained_data:
            time.sleep(TR_REQ_TIME_INTERVAL)
            # 이전 tr에서 마지막으로 저장한 날짜를 셋팅함. (self.lasted_date)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "일자", self.lasted_date)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.jongmokCode.text())
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "금액수량구분", "2")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "단위구분", "1")
            self._comm_rq_data("opt10059_req", "opt10059", 2, "0796")

    def _comm_rq_data(self, rqname, trcode, next, screen_no):

        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen_no)
        # 이벤트 루프 만들기
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

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
                self.rowDataTabWid.dataTable.setItem(crrOfRow, colidx,QTableWidgetItem(getdata))
                if colidx not in [0, 17]:
                    self._set_cell_style(crrOfRow, colidx, self.rowDataTabWid.dataTable.item(crrOfRow, colidx).text())
                colidx += 1
                self.lasted_date = self._comm_get_data(trcode, "", rqname, i, '일자')

            rowdatas.append(one_row_arr)

        if self.remained_data == False:
            print('data load ended.')
            self._make_sugup_data()
        else:
            print('.', end='')

    # 데이터 가져오기 함수
    def _comm_get_data(self, code, real_type, field_name, index, item_name):
        ret = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", code, real_type, field_name, index, item_name)
        return ret.strip()

    # 멀티데이터 크기 가져오기
    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    # 컬럼 폰트 스타일 설정
    def _set_cell_style(self, row, col, value):
        if value != '0':
            if value[:1] == '+' or value[:1] != '-':
                self.rowDataTabWid.dataTable.item(row, col).setForeground(QColor(255, 0 ,0))
            if value[:1] == '-':
                self.rowDataTabWid.dataTable.item(row, col).setForeground(QColor(0, 0, 255))

    # 수급 데이터 만들기
    def _make_sugup_data(self):
        print('수급데이터 만들기함수가 실행되었습니다')
        # 로우데이터 데이터를 ndarray객체로 만든다. ndarray객체가 수치계산에 더 빠르게 때문이다.
        self.np_row_data = np.array(rowdatas)
        self.np_sugup_data = np.zeros((self.np_row_data[:, 1].size, 87), dtype=int)

        # ['일자', '현재가', '전일대비', '등락율', '개인', '외국인', '기관계', '금융투자', '보험', '투신',
        # '기타금융', '은행', '연기금등', '사모펀드', '국가', '기타법인', '내외국인', '거래량']
        # 데이터역정렬 (역순계산 때문에)
        self.np_row_data = np.flipud(self.np_row_data)

        nd_gaein_data = np.zeros(5, dtype=int)      # 누적, 저점, 매집수, 매집고점, 분산비율
        for i in range(self.np_row_data.shape[0]):
            for j in range(self.np_row_data.shape[1]):
                if j == 0: self.np_sugup_data[i, j] = self.np_row_data[i, j]         # 일자
                if j == 1: self.np_sugup_data[i, j] = self.np_row_data[i, j]         # 현재가

                ## 여기서부터는 함수로 가능
                ## 세력합 제외 각 주체별 개별데이터 생성.
                if j in [2, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65] : self._make_sugup_part_data(j, i)

            nd_gaein_data[0] += int(self.np_row_data[i, 4])

        print(self.np_sugup_data)

    # 수급 주체별 데이터 generator
    def _make_sugup_part_data(self, fromidx, rowidx):
        for i in range(fromidx, fromidx+5):
            # 누적합
            if i == fromidx:
                if rowidx == 0:      self.np_sugup_data[rowidx, fromidx] = self.np_row_data[rowidx, 4]
                elif rowidx > 0:     self.np_sugup_data[rowidx, fromidx] = self.np_sugup_data[rowidx-1, i] + int(self.np_row_data[rowidx, 4])
            # 최고저점
            if i == fromidx+1:
                if rowidx == 0:      self.np_sugup_data[rowidx, fromidx]= self.np_row_data[i, 4]
                elif rowidx > 0:     self.np_sugup_data[rowidx, fromidx] = min(self.np_sugup_data[rowidx-1, i-1], self.np_sugup_data[rowidx, i-1])
            # 매집수량
            if i == fromidx+2: self.np_sugup_data[rowidx, fromidx] = self.np_sugup_data[rowidx, i-2] - self.np_sugup_data[i, i-3]
            # 매집고점
            if i == fromidx+3:
                if rowidx == 0:      self.np_sugup_data[rowidx, fromidx] = self.np_sugup_data[rowidx, i-1]
                elif rowidx > 0:     self.np_sugup_data[rowidx, fromidx] = max(self.np_sugup_data[rowidx-1, i], self.np_sugup_data[rowidx, i-1])
            # 분산비율
            if i == fromidx+4:
                if self.np_sugup_data[rowidx, i-2] == 0 or self.np_sugup_data[rowidx, i-1] == 0:
                    self.np_sugup_data[rowidx, fromidx] = 0
                else:
                    self.np_sugup_data[rowidx, fromidx] = (self.np_sugup_data[rowidx, i-2] / self.np_sugup_data[rowidx, i-1]) * 100

    # 엑셀파일로 데이터 저장
    def savefile(self):
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet("sheet 1", cell_overwrite_ok=True)
        self.add2(sheet)
        wbk.save("/Users/pconn/Desktop/export.xlsx")

    def add2(self, sheet):
        for currentColumn in range(self.rowDataTabWid.dataTable.columnCount()):
            for currentRow in range(self.rowDataTabWid.dataTable.rowCount()):
                try:
                    teext = str(self.rowDataTabWid.dataTable.item(currentRow, currentColumn).text())
                    sheet.write(currentRow, currentColumn, teext)
                except AttributeError:
                    pass

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
        self.sugupHeaders = ['거래량', '개인', '세력합', '외국인', '금융투자', '보험', '투신', '기타금융', '은행'
                                    , '연기금', '사모펀드', '국가', '기타법인', '내외국인']

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
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
