import os
import sys
import time
import json
import pandas as pd
import pyautogui as m
import pyperclip
import ctypes

import convert
# import main
# https://hoy.kr/7KlIe
m.PAUSE = 0.05
m.FAILSAFE = True

MsgBox = ctypes.windll.user32.MessageBoxW
# MsgBox(None, 'Hello', 'Window title', 0)

def get_xy(img):
    try:
        pos_img = m.locateOnScreen(img)
        x, y = m.center(pos_img)
        return (x, y)
    except Exception as e:
        pass

img_path = 'DZ.PNG'
dz_A = None
# dz_A = get_xy(img_path)
# _x, _y = dz_A
# dz_ilban = (_x + 100, _y + 223)
# print(dz_A)

class Duzon():
    def __init__(self):
        self.fromNzCode = convert.fromNzCode
        self.toDzCode = convert.toDzCode
        self.Dz_Card = convert.Dz_Card
        self.접대비계정 = convert.접대비계정
        self.check_xl = None
        self.df = None
        self.stop = True

    def ilban_xl_to_df(self, full_fn=None):

        if full_fn != None:
            try:
                # _df = pd.read_excel(full_fn, sheet_name='Sheet1', header=9)
                _df = pd.read_excel(full_fn, sheet_name=0, header=9)           
                df = _df.iloc[:, 0:9]
                self.check_xl = '년도월일' in df.columns[0]

                col = ['년월일', '구분', '계정코드', '계정과목', '거래처코드', '거래처명', '적요코드', '적요', '금액']
                df.columns = col
                df['년'] = df['년월일'].astype(str).apply(lambda x: x[:4])
                df['월'] = df['년월일'].astype(str).apply(lambda x: x[4:6])
                df['일'] = df['년월일'].astype(str).apply(lambda x: x[6:8])
                # data 전처리
                df['거래처코드'] = df['거래처코드'].fillna(0).astype(int).astype(str).apply(lambda x: x.zfill(5))
                df['적요코드'] = df['적요코드'].fillna(0).astype(int).astype(str).apply(lambda x: x.zfill(2))

                # 뉴젠계정코드 더존계정코드로  회전코드 피해가기 https://hoy.kr/x0dR9
                # df['계정코드'] = df['계정코드'].replace(fromNzCode,toDzCode)  # 회전코드 피해가기 https://hoy.kr/x0dR9
                self.df_idx_lst = list(df.index)
                self.Tatal_Line = len(self.df_idx_lst)
                for idx in self.df_idx_lst:
                    if df['계정코드'][idx] in self.fromNzCode:
                        idx_fromNzCode = self.fromNzCode.index(df['계정코드'][idx])
                        df['계정코드'][idx] = self.toDzCode[idx_fromNzCode]

                self.df = df
                # return self.check_xl, self.df
            except:
                pass

        return self.check_xl, self.df

    def input_dz(self, data):
        time.sleep(0.05)

        월 = data['월'].strip()
        일 = data['일'].strip()
        구분 = data['구분'].strip()
        계정코드 = data['계정코드'].strip()
        계정과목 = data['계정과목'].strip()
        거래처코드 = data['거래처코드'].strip()
        거래처명 = data['거래처명'].strip()
        금액 = data['금액'].strip()
        적요코드 = data['적요코드'].strip()
        적요 = data['적요'].strip()

        # m.typewrite(월)
        # m.typewrite(일)
        # m.typewrite(구분)
        # m.typewrite(계정코드)

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
        
        return True

        
        
