#!/usr/bin/env python
# coding: utf-8

# 예제 내용
# * QTreeWidget을 사용하여 아이템을 표시

__author__ = "Deokyu Lim <hong18s@gmail.com>"

import os
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtWidgets import QLabel, QProgressBar
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QBoxLayout, QFormLayout, QHBoxLayout
from PyQt5.QtGui import QIcon, QRegExpValidator, QDoubleValidator, QIntValidator, QFont, QPixmap
from PyQt5.QtCore import Qt, QRegExp, QThread, QWaitCondition, QMutex, pyqtSignal, pyqtSlot

from duzon import Duzon
import convert
# Total_Line = 300

class Thread(QThread):
    """
    단순히 0부터 100까지 카운트만 하는 쓰레드
    값이 변경되면 그 값을 change_value 시그널에 값을 emit 한다.
    """
    # 사용자 정의 시그널 선언
    change_value = pyqtSignal(int)

    def __init__(self):
        QThread.__init__(self)
        self.cond = QWaitCondition()
        self.mutex = QMutex()
        self.cnt = 0
        self.total_line = None
        self._status = False
        
    def __del__(self):
        self.wait()

    def run(self):
        while self.cnt < self.total_line:
            self.mutex.lock()

            if not self._status:
                self.cond.wait(self.mutex)

            # if 100 == self.cnt:
            #     self.cnt = 0
            self.cnt += 1
            self.change_value.emit(self.cnt)
            self.msleep(10)  # ※주의 QThread에서 제공하는 sleep을 사용

            self.mutex.unlock()

    def toggle_status(self):
        self._status = not self._status
        if self._status:
            self.cond.wakeAll()

    @property
    def status(self):
        return self._status


class Form(QWidget):
    def __init__(self):
        QWidget.__init__(self, flags=Qt.Widget)
        self.setWindowTitle("더존 일반전표 입력")

        self.th = Thread()
        # self.th.start()
                
        self.lb_path = QLabel("Excel:")
        self.pb_xl = QPushButton("엑셀선택")        
        self.le_start = QLineEdit()  # 시작년얼일
        self.le_end = QLineEdit()  # 종료년월일
        self.lb_cnt = QLabel()    # 총입력라인수
        self.lb_time = QLabel("00:00:00...")    # 총입력라인수
        self.pb_run = QPushButton("입력시작")
        self.pgb = QProgressBar()
        
        self.le_xy = QLineEdit()  # 일반전표입력좌표

        self.le_xy.setPlaceholderText("x좌표,y좌표")

        flo = QFormLayout()

        hbox1 = QHBoxLayout()
        flo1 = QFormLayout()
        flo2 = QFormLayout()
        flo1.addRow('시작일', self.le_start)
        flo2.addRow('종료일', self.le_end)
        hbox1.addLayout(flo1)
        hbox1.addLayout(flo2)

        hbox2 = QHBoxLayout()
        hbox2.addWidget(self.pb_xl)
        hbox2.addWidget(self.pb_run)

        flo.addRow(hbox1)
        flo.addRow(hbox2)
        flo.addRow(self.lb_path)

        hbox3 = QHBoxLayout()
        flo3 = QFormLayout()
        flo4 = QFormLayout()
        flo3.addRow('라인수:', self.lb_cnt)
        flo4.addRow('예상소요시간:', self.lb_time)
        hbox3.addLayout(flo3)
        hbox3.addLayout(flo4)

        flo.addRow(hbox3)
        flo.addRow(self.pgb)
        flo.addRow('일반전표입력 좌표', self.le_xy)

        # self.lb_test = QLabel()
        # flo.addRow('lb_test ', self.lb_test)


        # 입력제한 http://bitly.kr/wmonM2
        self.le_start.setMaxLength(8)
        self.le_end.setMaxLength(8)
        self.le_start.setInputMask("0000-00-00")
        self.le_end.setInputMask("0000-00-00")

        self.le_start.setStyleSheet(
            """QLineEdit {  width:55px; height:20px }""")   # ; width:30px; height:30px
        self.le_end.setStyleSheet(
            """QLineEdit { width:55px; height:20px }""")
        self.le_xy.setStyleSheet(
            """QLineEdit { width:40px; height:20px }""")
        self.lb_path.setStyleSheet(
            """QLabel { color: blue; font: bold; }""")    # font: 10pt}""")

        self.setLayout(flo)

        # 시그널 슬롯 연결
        self.pb_xl.clicked.connect(self.get_file_name)
        
        self.pb_run.clicked.connect(self.input_start)
        self.th.change_value.connect(self.pgb.setValue)
        # self.th.change_value.connect(self.Pause)

    @pyqtSlot()
    def get_file_name(self):
        self.pgb.reset()
        self.th.terminate()
        self.th.exit()
        
        filename = QFileDialog.getOpenFileName(
            caption="FILE DIALOG", directory=os.path.join(os.path.expanduser('~'), 'Desktop'), filter="*.xls*")
        self.full_fn = filename[0]
        fn = os.path.basename(self.full_fn)  # https://itmining.tistory.com/122
        self.lb_path.setText(fn)

        dz = Duzon(self.full_fn)
        df = dz.status()
        self.th.total_line = len(df)
        self.lb_cnt.setText(str(self.th.total_line))
        hms = convert.sec_to_hms(int(self.th.total_line * 1.56))
        if len(hms) == 3:
            hour, min, sec = hms[0], hms[1], hms[2]
            msg = f'{hour}:{min}:{sec}'
        elif len(hms) == 2:
            min, sec = hms[0], hms[1]
            if min < 10:
                min = str(min).zfill(2)
            msg = f'00:{min}:{sec}'
        elif len(hms) == 1:
            sec = hms[0]
            msg = f'00:00:{sec}'
        self.lb_time.setText(msg)
         
        # self.th.total_line = 1000
        df.to_excel('c:/zz/ttt.xlsx')
           
    @pyqtSlot()
    def input_start(self):        
        self.th.start()
        self.pgb.setMaximum(self.th.total_line)
        # 쓰레드 연결
        self.th.toggle_status()
        self.pb_run.setText({True: "입력중지", False: "입력시작"}[self.th.status])

    # @pyqtSlot(int)
    # def Pause(self, i):
        # self.pgb.setValue(i)
       
        


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    form = Form()
    form.show()
    exit(app.exec_())
