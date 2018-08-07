# 차트 탭 구성
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

class ChartTabWid(QWidget):
    def __init__(self, parent):
        self.layout = QVBoxLayout(self)

        # 탭 스크린 초기화
        self.chartTabs = QTabWidget()
        self.chartTab1 = QWidget()
        self.chartTab2 = QWidget()
        self.chartTab3 = QWidget()

        # 탭 추가
        self.chartTabs.addTab(self.chartTab1, "매집현황")
        self.chartTabs.addTab(self.chartTab2, "분산비율")
        self.chartTabs.addTab(self.chartTab3, "투자자추이")

        # 레이아웃 바인딩
        self.layout.addWidget(self.chartTabs)
        self.setLayout(self.layout)