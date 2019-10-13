import os
import sys
import time
import json
import pandas as pd
import pyautogui as m
import pyperclip

import convert

class Duzon():
    def __init__(self, full_fn=None):
        self.fromNzCode = convert.fromNzCode
        self.toDzCode = convert.toDzCode
        self.Dz_Card = convert.Dz_Card
        self.접대비계정 = convert.접대비계정
        self.check_xl = None
        self.df = None

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
            except:
                pass

    def status(self):
        return self.check_xl, self.df

        
