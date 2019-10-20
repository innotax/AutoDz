#!/usr/bin/env python
# coding: utf-8

# 예제 내용
# * QTreeWidget을 사용하여 아이템을 표시

__author__ = "Deokyu Lim <hong18s@gmail.com>"

import os
import sys
import json
import timeit
import pyautogui as m
import pyperclip

from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QLineEdit, QDateEdit, QComboBox
from PyQt5.QtWidgets import QLabel, QProgressBar
from PyQt5.QtWidgets import QApplication, QDesktopWidget, QMessageBox
from PyQt5.QtWidgets import QBoxLayout, QFormLayout, QHBoxLayout, QGroupBox
from PyQt5.QtGui import QIcon, QRegExpValidator, QDoubleValidator, QIntValidator, QFont, QPixmap
from PyQt5.QtCore import Qt, QDate, QRegExp, QThread, QWaitCondition, QMutex, pyqtSignal, pyqtSlot

# 상위폴더 내 파일 import  https://brownbears.tistory.com/296
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

import duzon 
import convert

# ===== Config =====
FULLPATH_SETUP_JSON = 'C:\zz\project\Automation\setup.json'
with open(FULLPATH_SETUP_JSON, encoding='utf-8') as fn:
    json_to_dict = json.load(fn)

SETUP_DIC = json_to_dict

m.PAUSE = 0.05
m.FAILSAFE = True
## ==================

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
        self.check_term = SETUP_DIC['Constant']['thread_check_interval']
        self.df = None
        self.total_line = None
        self.dz_A = None
        self.Dz_Card = SETUP_DIC['DzAccCode']['Card']
        self.접대비계정 = SETUP_DIC['DzAccCode']['접대비계정']

        self._status = False
        self.dz = duzon.Duzon()
        
    def __del__(self):
        self.wait()

    def set_velocity(self, dz_velocity, th_velocoty):
        self.dz_v = dz_velocity
        self.th_v = th_velocoty
        m.PAUSE = self.dz_v
        print("="*10)

    def run(self):
        try:
            print(img_path)
            self.dz_A = get_xy(img_path)
            print(dz_A)
        except:
            pass

        if self.dz_A !=None:
            m.click(self.dz_A)
            # _x, _y = self.dz_A
            # dz_ilban = (_x + 100, _y + 223)
            # m.click(dz_ilban)
            m.hotkey('ctrl', 'w')
            dz_ilban_shortcut = SETUP_DIC['DzShotcut']['일반전표입력']
            m.typewrite(dz_ilban_shortcut)
            m.press('enter')
            self.msleep(100)
            print("start"*10)
            start_month = str(self.df['년월일'].min())[4:6]
            end_month = str(self.df['년월일'].max())[4:6]
            m.typewrite(start_month)
            m.typewrite(end_month)
            self.msleep(100)
            df_idx_lst = list(self.df.index)
            self.msleep(1000)

            for self.cnt, idx in enumerate(df_idx_lst[:]):
                # goto = False
                self.mutex.lock()

                if not self._status:
                    self.cond.wait(self.mutex)
                # send_data = self.make_send_dict(idx)
                self.msleep(self.th_v)
                # print('idx',idx)
                # self.dz.input_dz(send_data)
                self.input_dz(idx)
                self.msleep(2000)
                
                self.change_value.emit(self.cnt)
                if self.cnt % self.check_term ==0:
                    self.change_term.emit()

                self.cnt += 1            
                # self.msleep(self.th_v)  # ※주의 QThread에서 제공하는 sleep을 사용
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

    def input_dz(self, idx):
        # time.sleep(0.05)
        # self.msleep(self.th_v)

        월 = str(self.df['월'][idx]).strip()
        일 = str(self.df['일'][idx]).strip()
        구분 = str(self.df['구분'][idx]).strip()
        계정코드 = str(self.df['계정코드'][idx]).strip()
        계정과목 = self.df['계정과목'][idx].strip()
        거래처코드 = str(self.df['거래처코드'][idx]).strip()
        거래처명 = str(self.df['거래처명'][idx]).strip()
        금액 = str(self.df['금액'][idx]).strip()
        적요코드 = str(self.df['적요코드'][idx]).strip()
        적요 = str(self.df['적요'][idx]).strip()

        # m.typewrite(월)
        # m.typewrite(일)
        # m.typewrite(구분)
        # m.typewrite(계정코드)
        # self.msleep(self.th_v)

        pyperclip.copy(월)
        m.hotkey('ctrl', 'v')
        m.press('enter')
        pyperclip.copy(일)
        m.hotkey('ctrl', 'v')
        m.press('enter')
        pyperclip.copy(구분)
        m.hotkey('ctrl', 'v')
        m.press('enter')
        pyperclip.copy(계정코드)
        m.hotkey('ctrl', 'v')
        m.press('enter')
        # time.sleep(0.05)

        if 거래처코드 != "00000":
            pyperclip.copy(거래처코드)
            m.hotkey('ctrl', 'v')
            m.press('enter')
        else:
            m.press('enter')

            pyperclip.copy(거래처명)
            m.hotkey('ctrl', 'v')
            m.press('enter')

        pyperclip.copy(금액)
        m.hotkey('ctrl', 'v')
        m.press('enter')

        if 계정코드 in self.접대비계정:
            적요코드 = '02'
            m.press('enter')

        elif (int(구분) in [1, 3] and int(금액) > 30000
                and int(계정코드) in self.Dz_Card):
            m.press('enter')
            pyperclip.copy(적요)
            m.hotkey('ctrl', 'v')
            m.press('enter')
            m.press('enter')
            m.press('enter')

        else:
            m.press('enter')
            pyperclip.copy(적요)
            m.hotkey('ctrl', 'v')
            m.press('enter')

        self.msleep(self.th_v)

    def toggle_status(self):
        self._status = not self._status
        if self._status:
            self.cond.wakeAll()

    @property
    def status(self):
        return self._status


class MainGUI(QWidget):
    def __init__(self):
        QWidget.__init__(self, flags=Qt.Widget)
        rect = QDesktopWidget().availableGeometry()   # 작업표시줄 제외한 화면크기 반환
        max_x = rect.width()
        max_y = rect.height()
        width = SETUP_DIC['Constant']['main_gui_width']
        height = SETUP_DIC['Constant']['main_gui_height']
        left = max_x - width
        top = max_y - height
        self.setGeometry(left, top, width, height)
        self.setWindowTitle("더존 일반전표 입력")

        self.setup_dic = SETUP_DIC
                
        self.lb_path = QLabel("Excel:")
        self.pb_xl = QPushButton("엑셀선택")        
        self.dateed_start = QDateEdit()  # 시작년얼일
        self.dateed_start.setDate(QDate.currentDate())
        self.dateed_start.setCalendarPopup(True)
        self.dateed_end = QDateEdit()  # 종료년월일
        self.dateed_end.setDate(QDate.currentDate())
        self.dateed_end.setCalendarPopup(True)

        self.lb_cnt = QLabel()    # 총입력라인수
        self.lb_time = QLabel("00:00:00...")    # 예상소요시간
        self.pb_term = QPushButton("기간재설정")
        self.pb_run = QPushButton("입력시작")
        self.pb_help = QPushButton()
        self.pb_help.setIcon(QIcon('data/help.ico'))
        self.pb_help.setFixedWidth(30)
        self.pb_help.setStyleSheet(
            """QPushButton { border-image: url(:data/help.ico); width:20px; height:20px }""")

        self.pgb = QProgressBar()
        
        self.cb_dz = QComboBox()  # 스레드속도
        self.cb_th = QComboBox()  # 스레드속도

        dz_velocity_level = SETUP_DIC['velocity']['pyautogui']['level']
        dz_velocity_items = list(dz_velocity_level.keys())
        self.cb_dz.addItems(dz_velocity_items)
        self.dz_velocity_default = SETUP_DIC['velocity']['pyautogui']['default']
        for k, v in dz_velocity_level.items():
            if v == self.dz_velocity_default:
                dz_level = k
        self.cb_dz.setCurrentIndex(dz_velocity_items.index(dz_level))

        th_velocity_level = SETUP_DIC['velocity']['QThread']['level']
        th_velocity_items = list(th_velocity_level.keys())
        self.cb_th.addItems(th_velocity_items)
        self.th_velocity_default = SETUP_DIC['velocity']['QThread']['default']
        for k, v in th_velocity_level.items():
            if v == self.th_velocity_default:
                th_level = k
        self.cb_th.setCurrentIndex(th_velocity_items.index(th_level))

        self.le_dz_shortcut = QLineEdit("단축키")
        self.le_dz_shortcut.setStyleSheet(
            """QLineEdit { width:20px; height:20px }""")
        self.pb_save = QPushButton()
        # QIcon width # pyinstaller image err solution
        self.pb_save.setIcon(QIcon('data/save.ico'))
        self.pb_save.setFixedWidth(30)
        self.pb_save.setStyleSheet(   
            """QPushButton { border-image: url(:data/save.ico); width:20px; height:20px }""")
        
        # self.le_xy.setPlaceholderText("x좌표,y좌표")

        flo = QFormLayout()

        hbox1 = QHBoxLayout()
        flo1 = QFormLayout()
        flo2 = QFormLayout()
        flo1.addRow('시작일', self.dateed_start)
        flo2.addRow('종료일', self.dateed_end)
        hbox1.addLayout(flo1)
        hbox1.addLayout(flo2)

        hbox2 = QHBoxLayout()
        hbox2.addWidget(self.pb_xl)
        hbox2.addWidget(self.pb_term)
        hbox2.addWidget(self.pb_run)
        hbox2.addWidget(self.pb_help)

        hbox3 = QHBoxLayout()
        flo3 = QFormLayout()
        flo4 = QFormLayout()
        flo3.addRow('라인수:', self.lb_cnt)
        flo4.addRow('남은시간:', self.lb_time)
        hbox3.addLayout(flo3)
        hbox3.addLayout(flo4)

        hbox4 = QHBoxLayout()
        hbox4.addWidget(self.cb_dz)
        hbox4.addWidget(self.cb_th)
        hbox4.addWidget(self.le_dz_shortcut)
        hbox4.addWidget(self.pb_save)

        gbx = QGroupBox("속도:1.더존,2.쓰레드 / 전표입력단축키/저장")
        gbx.setLayout(hbox4)

        flo.addRow(hbox1)
        flo.addRow(hbox2)
        flo.addRow(self.lb_path)
        flo.addRow(hbox3)
        flo.addRow(self.pgb)
        flo.addRow(gbx)  # "속도(쓰레드,더존)/단축키",

        # self.le_xy.setStyleSheet(
        #     """QLineEdit { width:40px; height:20px }""")
        self.lb_path.setStyleSheet(
            """QLabel { color: blue; font: bold; }""")    # font: 10pt}""")
        self.pb_xl.setStyleSheet(
            """QPushButton { background-color: #c5d1e9; color: blue; font: bold }""")  # ffff00

        self.setLayout(flo)

class Main(MainGUI):
    def __init__(self):
        super().__init__()

        self.th = Thread()
        # self.th.start()
        self.dz = duzon.Duzon()
        self.df = None
        self.check_term_flag = False
        self.count_pause = 0
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

            start_YMD = str(self.df['년월일'].min())
            s_yy, s_mm, s_dd = int(start_YMD[:4]), int(start_YMD[4:6]), int(start_YMD[6:8])
            self.dateed_start.setDate(QDate(s_yy, s_mm, s_dd))
            end_YMD = str(self.df['년월일'].max())
            e_yy, e_mm, e_dd = int(end_YMD[:4]), int(end_YMD[4:6]), int(end_YMD[6:8])
            self.dateed_end.setDate(QDate(e_yy, e_mm, e_dd))

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
                self.dateed_start.setStyleSheet(
                    """QDateEdit { background-color: #c5d1e9; color: blue;}""")
                self.dateed_end.setStyleSheet(
                    """QDateEdit { background-color: #c5d1e9; color: blue; }""")
                self.pb_term.setStyleSheet(
                    """QPushButton { background-color: #c5d1e9; color: blue; font: bold }""")
            else:
                self.pb_run.setStyleSheet(
                    """QPushButton { background-color: #c5d1e9; color: blue; font: bold }""")

                pass
            
            self.df.to_excel('c:/zz/ttt.xlsx')

    @pyqtSlot()
    def set_term(self):
        if len(self.df)>0:
            start = "".join(str(self.dateed_start.date().toPyDate()).split("-"))
            end = "".join(str(self.dateed_end.date().toPyDate()).split("-"))
            start_day = int(start)
            end_day = int(end)
            self.df = self.df[(self.df['년월일'] >= start_day) & (self.df['년월일'] <= end_day)]
            self.th.df = self.df
            self.th.total_line = len(self.df)
            self.lb_cnt.setText(str(self.th.total_line))

            hms_msg = convert.sec_to_hms_msg(int(self.th.total_line * 1.56))
            self.lb_time.setText(hms_msg)
            self.pb_run.setStyleSheet(
                """QPushButton { background-color: #c5d1e9; color: blue; font: bold }""")

            self.df.to_excel('c:/zz/ttt.xlsx')
        else:
            print("E"*20)
            pass
        
            
    @pyqtSlot()
    def input_start(self):
        msgbox_title = "거래처등록(일반,금융,카드)"
        msg = """더존 거래처코드와 수지라 거래처코드가<br>
                일치하지 않으면 오류가 발생합니다 !!! <br>
                꼭 6.붙여넣기목록 > 학습로직 에서 <br>
                기초코드등록 후 작업을 시작하세요 !<br>
                YES(입력시작)"""
        msgTF = MsgBoxTF(msgbox_title, msg)
        TF = msgTF.initUI()
        if TF:
            if self.th.total_line > 0:
                print(self.th.total_line)
                print(self.dz_velocity_default, self.th_velocity_default)
                self.count_pause += 1
                self.th.set_velocity(
                    self.dz_velocity_default, self.th_velocity_default)
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
                    if self.count_pause > 1:
                        m.click(self.th.dz_A)
                        m.press('pagedown')
                else:
                    self.pb_run.setStyleSheet(
                        """QPushButton { background-color: #c5d1e9; color: blue; font: bold }""")
                    # m.click(self.th.dz_A)
                    # m.press('pagedown')
            else:
                pass
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
    main = Main()
    main.show()
    exit(app.exec_())
