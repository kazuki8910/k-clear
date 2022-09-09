#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 集客表のデータ取得
# クランROASシートの基幹データ更新


########################
# 
# 設定（変動）
# 
########################

# 該当月
applicable_month = 9

# 02シートのタブ名
name_02 = "22年9月"

# 更新するシート名
name_to = "まとめデータ_22年9月"




########################
# 
# モジュール
# 
########################

# 一般モジュール
import pandas as pd
import re
import pytz
import sys
import os
import csv
import time
from datetime import datetime as dt
from selenium.webdriver.common.by import By

# スプシ関連
from gspread_dataframe import set_with_dataframe

# 自作モジュール
import func
import conf

print("モジュールのインポート完了")



########################
# 
# 設定
# 
########################

# シートキー
key_02 = "1UvGMEFNqhRE-YfuYV1nwhKUb3YWCqRHRttwCGwXZO3g"

# 成果シートを更新するループの関数
def update_sheet(key_list, name_to, df_origin):
    for this_key in key_list:
        i=0
        while True:
            i=i+1
            if(i>10):
                sys.exit()

            try:
                asp_name = this_key[0]
                key_to = this_key[1]

                # ROAS管理シートにアクセス
                wb_to = func.connect_gspread(key_to)
                ws_to = wb_to.worksheet(name_to)

                # まとめデータを貼り付け
                set_with_dataframe(ws_to, df_origin)

                # タイムスタンプ
                cell = ws_to.range(1,1,1,1)
                cell[0].value = str(dt.now(pytz.timezone('Asia/Tokyo')))
                ws_to.update_cells(cell)
                
                print(asp_name)
                break

            except:
                print(asp_name + "更新失敗")




########################
# 
# クランの基幹タブ更新
# 
########################
########################
# 変数
########################
# 基幹システムのURL
url_kikan = "https://mensclear.com/users/reservations"
# ログイン情報
login_email = "cb140000@yahoo.co.jp" # メールアドレス
login_pass  = "muramura" # パスワード
# CSVファイルの保存先
csv_filepass = "data/reservations.csv"
# 実行月
this_month = dt.today().month

#########################
# 月をまたいだら実行しない
#########################
if(this_month == applicable_month):

    # クローム起動
    drive = func.start_chrome()

    #########################
    # 基幹システムにログイン
    #########################

    drive.get(url_kikan)
    time.sleep(3)

    # メールに入力
    elm_email = drive.find_element(By.ID, "user_email")
    elm_email.send_keys(login_email)
    time.sleep(1)

    # パスワードを入力
    elm_pass = drive.find_element(By.ID, "user_password")
    elm_pass.send_keys(login_pass)
    time.sleep(1)

    # ログインボタン押下
    elm_login_btn = drive.find_element(By.CLASS_NAME, "p-sessions__button")
    elm_login_btn.click()
    time.sleep(3)

    #########################
    # CSVダウンロード
    #########################

    # データ一覧を表示
    elm_data_index = drive.find_elements(By.CLASS_NAME, "l-dashboard__panel")[0]
    elm_data_index.click()
    time.sleep(3)

    # ダウンロードボタンをクリック
    elm_dl_btn = drive.find_element(By.NAME, "commit")
    elm_dl_btn.click()
    time.sleep(3)

    #########################
    # スプシにアップロード
    #########################

    # クラン ROAS管理シートキー
    sheet_key = "1GiD6IFZhjhVe-y-DvEUoT0et6Qt_mYBRKF_tWGR5_z8"

    # シートにアクセス
    wb = func.connect_gspread(sheet_key)

    # CSVをアップロード
    wb.values_update(
        '基幹_当月',
        params={'valueInputOption': 'USER_ENTERED'},
        body={'values': list(csv.reader(open(csv_filepass, encoding='cp932')))}
    )
    time.sleep(3)

    #########################
    # CSVファイルを削除
    #########################

    os.remove(csv_filepass)

    drive.quit()

    print("クランの基幹タブ更新完了")

else:
    print("月を跨いだので設定を更新してください")





########################
# 
# 成果確認シートの集客表タブ更新
# 
########################

key_list = [
    ["フォースリー 成果", "1p-NEAiD3uHUp-Gcvzi6bm8LF9NJYDQ5Ju24HwOS8UYU"],
    ["リンクエッジ 成果", "1XbV5cgwkc6VuXlCUOSviqE4IdkbpyDxqug88089POfE"],
    # ["サルクルー 成果", ""],
    ["FORCE 成果", "1SqjCFNE-gn_5feLbXx3iqnsWSR4e8RQtsfj3mkOlYTk"],
    # ["レントラ（FA） 成果", ""],
    ["セレス 成果", "1WqTDLkMwHbSu5mnNZgJRjShaXqM6LC3Fi9ItPJdwxdI"],
    # ["ブランディングエンジニア 成果", ""],
    # ["パフォテク 成果", ""],
    ["アレテコ 成果", "1d6mU8C7i3pFacOJrcgQfczZugHYe-B1pSjjYBDSWvyY"],
    # ["ブリーチ 成果", ""],
    ["FA 成果", "1L0DJRvcvu8x58nz0MWMB5Pu-JMrbk7D4RE3ep-ALioo"],
    ["ナハト（クラン） 成果", "14GhIRCLUQWB-oPX3unXwyaqErJEOPEPwaHuhOAZwFIU"],
    ["FORCE 成果", "1S9rQlhBj_VjpAvFRJXbfoDKzQigFX0T7OvM39nBpKFM"],
    # ["クリエイト 成果", ""],
    ["h1 成果", "1-eU5PeNBNXqIONfIRZxJErxItERh6DhVlrx-9IewPhI"],
    ["h3 成果", "1_JWsy811ICtYojOoVw_wm5lHLnpfXh0n8_r-zoIza3E"],
    ["GDN 成果", "1_JWsy811ICtYojOoVw_wm5lHLnpfXh0n8_r-zoIza3E"]
]


# 集客表元データのシートにアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)

# 元データ取得
df_origin = pd.DataFrame(ws_origin.get_all_values()) # 元データ取得
df_origin = df_origin[df_origin[4] != ""]   # ID空白除去


# ループにする
update_sheet(key_list, name_to, df_origin)

print("成果確認シート 完了")


# In[ ]:


########################
# 
# ROAS管理シートの集客表タブ更新
# 
########################

key_list = [
    ["ナハト(クラン) ROAS", "1GiD6IFZhjhVe-y-DvEUoT0et6Qt_mYBRKF_tWGR5_z8"]
]

# シート名
name_to = "集客表_当月"


# 集客表元データのシートにアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)

# 元データ取得
df_origin = pd.DataFrame(ws_origin.get_all_values()) # 元データ取得
df_origin = df_origin[df_origin[4] != ""]   # ID空白除去


# ループにする
update_sheet(key_list, name_to, df_origin)

print("ROAS確認シート 完了")

