import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import  *
from PyQt5.QAxContainer import *

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Kiwoom Login
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.dynamicCall("CommConnect()")
        self.kiwoom.OnReceiveTrData.connect(self._receive_tr_data)

        self.setWindowTitle("종목코드")
        self.setGeometry(0,0,1800,900)

        btn1 = QPushButton("조회", self)
        btn1.move(120, 10)
        btn1.clicked.connect(self.btn1_clicked)

        self.jongmokCode = QLineEdit(self)
        self.jongmokCode.move(10,10)


        ## 종목별투자자기관별요청 데이터조회 테이블
        ## 테이블 헤더설정
        self.headers = ['일자', '현재가', '전일대비', '등락율', '개인', '외국인', '기관계', '금융투자', '보험', '투신',
                   '기타금융', '은행', '연기금등', '사모펀드', '국가', '기타법인', '내외국인', '거래량']

        self.dataTable = QTableWidget(0, self.headers.__len__(), self)
        self.dataTable.setRowHeight(0, 10)
        self.dataTable.setGeometry(10, 700, 1780, 190)
        self.dataTable.setHorizontalHeaderLabels(self.headers)

        ##종목별투자자기관별요청 컬럼정보
        self.jmTabColItemInfo = ['일자', '현재가', '전일대비', '등락율', '개인투자자', '외국인투자자', '기관계', '금융투자',
                            '보험', '투신', '기타금융', '은행', '연기금등', '사모펀드', '국가', '기타법인', '내외국인',
                            '누적거래대금']

    def btn1_clicked(self):
        # 조회전 기존 데이터 리셋.
        self.dataTable.setRowCount(0)

        # 종목별투자자기관별요청

        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "일자", "20180713")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.jongmokCode.text())
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "금액수량구분", "2")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "단위구분", "1")
        returnCode = self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10059_req", "opt10059", 0, "0796")
        print(returnCode)

    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        if rqname == 'opt10059_req':
            data_cnt = self._get_repeat_cnt(trcode, rqname)
            print(data_cnt)
            for i in range(data_cnt):
                crrOfRow = self.dataTable.rowCount()
                self.dataTable.setRowCount(crrOfRow + 1)
                self.dataTable.setRowHeight(crrOfRow, 10)

                colidx = 0
                for j in self.jmTabColItemInfo:
                    self.dataTable.setItem(crrOfRow, colidx, QTableWidgetItem(self._comm_get_data(trcode, "", rqname, i, j)))
                    if colidx not in [0, 17]:
                       self._set_cell_style(crrOfRow, colidx, self.dataTable.item(crrOfRow, colidx).text())

                    colidx += 1

    ##데이터 가져오기 함수
    def _comm_get_data(self, code, real_type, field_name, index, item_name):
        ret = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", code, real_type, field_name, index, item_name)
        return ret.strip()

    ##멀티데이터 크기 가져오기
    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    ##컬럼 폰트 스타일 설정
    def _set_cell_style(self, row, col, value):
        if value != '0':
            if value[:1] == '+' or value[:1] != '-':
                self.dataTable.item(row, col).setForeground(QColor(255, 0 ,0))
            if value[:1] == '-':
                self.dataTable.item(row, col).setForeground(QColor(0, 0, 255))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()