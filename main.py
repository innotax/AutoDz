#!/usr/bin/env python
# coding: utf-8

# 예제 내용
# * QTreeWidget을 사용하여 아이템을 표시

__author__ = "Deokyu Lim <hong18s@gmail.com>"

import os
import sys
import timeit
import pyautogui as m

from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtWidgets import QLabel, QProgressBar
from PyQt5.QtWidgets import QApplication, QDesktopWidget, QMessageBox
from PyQt5.QtWidgets import QBoxLayout, QFormLayout, QHBoxLayout
from PyQt5.QtGui import QIcon, QRegExpValidator, QDoubleValidator, QIntValidator, QFont, QPixmap
from PyQt5.QtCore import Qt, QRegExp, QThread, QWaitCondition, QMutex, pyqtSignal, pyqtSlot

import duzon 
import convert


def get_xy(img):
    try:
        pos_img = m.locateOnScreen(img)
        x, y = m.center(pos_img)
        return (x, y)
    except Exception as e:
        pass


img_path = 'C:\zz\project\Automation\DZ.PNG'
# dz_A = None


class MsgBoxTF(QMessageBox):
    '''param : title="QMessageBox", msg=None
       return : True / False
    '''

    def __init__(self, title="QMessageBox", msg=None):
        super().__init__()
        self.title = title
        self.msg = msg

        rect = QDesktopWidget().availableGeometry()   # 작업표시줄 제외한 화면크기 반환
        max_x = rect.width()
        max_y = rect.height()

        self.width = 320
        self.height = 550
        self.left = max_x - self.width
        self.top = max_y - self.height
        
        # self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        buttonReply = QMessageBox.question(
            self, self.title, self.msg, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if buttonReply == QMessageBox.Yes:
            return True
        else:
            return False

class Thread(QThread):
    """
    단순히 0부터 100까지 카운트만 하는 쓰레드
    값이 변경되면 그 값을 change_term 시그널에 값을 emit 한다.
    """
    # 사용자 정의 시그널 선언
    change_value = pyqtSignal(int)
    change_term = pyqtSignal()

    def __init__(self):
        QThread.__init__(self)
        self.cond = QWaitCondition()
        self.mutex = QMutex()
        self.cnt = 0
        self.df = None
        self.total_line = None
        self.check_term = 10
        self.dz_A = None

        self._status = False
        self.dz = duzon.Duzon()
        
    def __del__(self):
        self.wait()

    def run(self):
        try:
            print(img_path)
            self.dz_A = get_xy(img_path)
            print(dz_A)
        except:
            pass

        if self.dz_A !=None:
            m.click(self.dz_A)
            _x, _y = self.dz_A
            dz_ilban = (_x + 100, _y + 223)
            m.click(dz_ilban)
            self.msleep(1000)
            print("start"*10)
            start_month = str(self.df['년월일'].min())[4:6]
            end_month = str(self.df['년월일'].max())[4:6]
            m.typewrite(start_month)
            m.typewrite(end_month)
            self.msleep(100)
            df_idx_lst = list(self.df.index)

            for self.cnt, idx in enumerate(df_idx_lst[:]):
                # goto = False
                self.mutex.lock()

                if not self._status:
                    self.cond.wait(self.mutex)
                send_data = self.make_send_dict(idx)
                self.msleep(100)
                # print('idx',idx)
                self.dz.input_dz(send_data)
                self.msleep(100)
                # result = self.dz.input_dz(send_data)
                # yield result
                # if result:
                self.change_value.emit(self.cnt)
                if self.cnt % self.check_term ==0:
                    self.change_term.emit()

                self.cnt += 1            
                self.msleep(50)  # ※주의 QThread에서 제공하는 sleep을 사용
                self.mutex.unlock()
    
    def make_send_dict(self, idx):
        data = dict()
        data['월'] = str(self.df['월'][idx])
        data['일'] = str(self.df['일'][idx])
        data['구분'] = str(self.df['구분'][idx])
        data['계정코드'] = str(self.df['계정코드'][idx])
        data['계정과목'] = self.df['계정과목'][idx]
        data['거래처코드'] = str(self.df['거래처코드'][idx])
        data['거래처명'] = str(self.df['거래처명'][idx])
        data['금액'] = str(self.df['금액'][idx])
        data['적요코드'] = str(self.df['적요코드'][idx])
        data['적요'] = str(self.df['적요'][idx])
        
        return data

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
        rect = QDesktopWidget().availableGeometry()   # 작업표시줄 제외한 화면크기 반환
        max_x = rect.width()
        max_y = rect.height()

        width, height = 320, 200
        left = max_x - width
        top = max_y - height
        self.setGeometry(left, top, width, height)
        self.setWindowTitle("더존 일반전표 입력")

        self.th = Thread()
        # self.th.start()
        self.dz = duzon.Duzon()
        self.df = None
        self.check_term_flag = False
                
        self.lb_path = QLabel("Excel:")
        self.pb_xl = QPushButton("엑셀선택")        
        self.le_start = QLineEdit()  # 시작년얼일
        self.le_end = QLineEdit()  # 종료년월일
        self.lb_cnt = QLabel()    # 총입력라인수
        self.lb_time = QLabel("00:00:00...")    # 예상소요시간
        self.pb_term = QPushButton("기간재설정")
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
        hbox2.addWidget(self.pb_term)
        hbox2.addWidget(self.pb_run)

        flo.addRow(hbox1)
        flo.addRow(hbox2)
        flo.addRow(self.lb_path)

        hbox3 = QHBoxLayout()
        flo3 = QFormLayout()
        flo4 = QFormLayout()
        flo3.addRow('라인수:', self.lb_cnt)
        flo4.addRow('남은시간:', self.lb_time)
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
        self.le_start.setInputMask("9999-99-99")
        self.le_end.setInputMask("0000-00-00")

        self.le_start.setStyleSheet(
            """QLineEdit {  width:55px; height:20px }""")   # ; width:30px; height:30px
        self.le_end.setStyleSheet(
            """QLineEdit { width:55px; height:20px }""")
        self.le_xy.setStyleSheet(
            """QLineEdit { width:40px; height:20px }""")
        self.lb_path.setStyleSheet(
            """QLabel { color: blue; font: bold; }""")    # font: 10pt}""")
        self.pb_xl.setStyleSheet(
            """QPushButton { background-color: #ffff00; color: blue; font: bold }""")

        self.setLayout(flo)

        # 시그널 슬롯 연결
        self.pb_xl.clicked.connect(self.get_file_name)
        self.pb_term.clicked.connect(self.set_term)
        
        self.pb_run.clicked.connect(self.input_start)
        self.th.change_value.connect(self.pgb.setValue)
        self.th.change_term.connect(self.receive_change_term)

    @pyqtSlot()
    def get_file_name(self):
        self.pgb.reset()
        self.th.terminate()
        self.th.exit()
        self.lb_path.setText("Excel:")
        self.lb_cnt.clear()
        self.lb_time.setText("00:00:00...")
        self.le_start.clear()
        self.le_end.clear()
        
        filename = QFileDialog.getOpenFileName(
            caption="FILE DIALOG", directory=os.path.join(os.path.expanduser('~'), 'Desktop'), filter="*.xls*")
        self.full_fn = filename[0]
        fn = os.path.basename(self.full_fn)  # https://itmining.tistory.com/122
        
        check_xl, self.df = self.dz.ilban_xl_to_df(self.full_fn)
        if check_xl:
            self.lb_path.setText(fn)
            self.th.df = self.df
            self.th.total_line = len(self.df)
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

            # self.le_start.setInputMask("")
            # self.le_start.setPlaceholderText(str(self.df['년월일'].min()))
            self.le_start.setText(str(self.df['년월일'].min()))
            # self.le_start.setInputMask("0000-00-00")
            self.le_end.setText(str(self.df['년월일'].max()))

            self.pb_xl.setStyleSheet(
                """QPushButton { background-color: #e1e1e1;}""")
            self.pb_term.setStyleSheet(
                """QPushButton { background-color: #e1e1e1;}""")
            self.pb_run.setStyleSheet(
                """QPushButton { background-color: #e1e1e1;}""")
            self.lb_cnt.setStyleSheet(
                """QLabel { color: blue; font: bold; }""")
            self.lb_time.setStyleSheet(
                """QLabel { color: blue; font: bold; }""")


            msgbox_title = "입력기간 설정"
            msg = """입력 시작년월일과 종료년월일을<br>
                     재설정 하시겠습니까 ? <br>
                     시작일과 종료일에 기간입력 후 <br>
                     기간재설정 버튼을 누르세요!!!"""
            msgTF = MsgBoxTF(msgbox_title, msg)
            TF = msgTF.initUI()
            if TF:
                self.le_start.setStyleSheet(
                    """QLineEdit { background-color: #ffff00; color: blue;}""")
                self.le_end.setStyleSheet(
                    """QLineEdit { background-color: #ffff00; color: blue; }""")
                self.pb_term.setStyleSheet(
                    """QPushButton { background-color: #ffff00; color: blue; font: bold }""")
            else:
                self.pb_run.setStyleSheet(
                    """QPushButton { background-color: #ffff00; color: blue; font: bold }""")

                pass
            
            self.df.to_excel('c:/zz/ttt.xlsx')

    @pyqtSlot()
    def set_term(self):
        if len(self.df)>0:
            start = self.le_start.text().replace('-','')
            end = self.le_end.text().replace('-','')
            start_day = int(start)
            end_day = int(end)
            self.df = self.df[(self.df['년월일'] >= start_day) & (self.df['년월일'] <= end_day)]
            self.th.df = self.df
            self.th.total_line = len(self.df)
            self.lb_cnt.setText(str(self.th.total_line))

            hms_msg = convert.sec_to_hms_msg(int(self.th.total_line * 1.56))
            self.lb_time.setText(hms_msg)
            self.pb_run.setStyleSheet(
                """QPushButton { background-color: #ffff00; color: blue; font: bold }""")

            self.df.to_excel('c:/zz/ttt.xlsx')
        else:
            print("E"*20)
            pass
        
            
    @pyqtSlot()
    def input_start(self):
        if self.th.total_line > 0:
            self.th.start()
            self.pgb.setMaximum(self.th.total_line -1)
            # 쓰레드 연결
            self.th.toggle_status()
            self.pb_run.setText({True: "입력중지", False: "입력시작"}[self.th.status])
            # self.dz.input_dz(self.th.status)
            if self.th.status:
                self.pb_term.setStyleSheet(
                    """QPushButton { background-color: #e1e1e1;}""")
                self.pb_run.setStyleSheet(
                    """QPushButton { background-color: #ff0000; color: yellow; font: bold }""")
                
                m.click(self.th.dz_A)
                m.press('pagedown')
            else:
                self.pb_run.setStyleSheet(
                    """QPushButton { background-color: #ffff00; color: blue; font: bold }""")  
                # m.click(self.th.dz_A)
                # m.press('pagedown')
        else:
            pass

        
    @pyqtSlot()
    def receive_change_term(self):
        self.check_term_flag = not self.check_term_flag

        if self.check_term_flag:
            self.start = timeit.default_timer()            
        else:
            self.end = timeit.default_timer()
            per_sec = (self.end - self.start) / self.th.check_term
            expect_sec = int((self.th.total_line - self.th.cnt)*per_sec)
            hms_msg = convert.sec_to_hms_msg(expect_sec)
            self.lb_time.setText(hms_msg)
            line_change_msg = f'{self.th.cnt}/{self.th.total_line}'
            self.lb_cnt.setText(line_change_msg)
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    form = Form()
    form.show()
    exit(app.exec_())
