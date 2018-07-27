##### 공통적으로 사용할 함수들 모음

class RqFunc:

    ## commRqData -- tr요청
    def _comm_rq_data(self, rqname, trcode, next, screenNo):
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screenNo)