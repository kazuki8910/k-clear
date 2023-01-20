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
applicable_month = 1

# 02シートのタブ名
name_02 = "23年1月"

# 更新するシート名
name_to = "まとめデータ_23年1月"

# 02シートのキー
key_02 = "1rv1gIrocxGyeD71mhr3_gKfMbD24foEoWCOwVX2eR-8"

# ROAS管理シートのキー
key_list_roas = [
    ["フォースリー", "1HLf2S8f485cJZ5lbQJEUbhTyx9I5t94Cj9jp29JUlcA"],
    ["FA", "1188aA2qg-MAAyfLaVaJj_LLfK7EdW90I5qQhdwUAZa8"],
    ["ナハト", "157R3oGs7uYgtkm7AZo59dnwuIeIy8FRILCAJEWtR7Ls"],
    ["FORCE", "1qzpsEhX5S2o5fRpFMZkkxZygxWbwegoTzckMxTzTPGI"],
]

# 成果確認シートのキー
key_list_seika = [
    # ["フォースリー 成果", "1sf6uWbPFSjmTkyagLZD_DCSWR1lGfij3V9TjfXxbnqg"],
    # ["リンクエッジ 成果", "1KpSLnbKHHutWizb9qldIJl7-Q943PkAolubjAXED8oU"],
    # ["サルクルー 成果", ""],
    # ["FORCE(レントラ) 成果", "1v-o4RwUEUVvLiPnyo4nhWpnb9n826SlhQovkHIjljGw"],
    # ["レントラ（FA） 成果", ""],
    # ["セレス 成果", "153vDemv8gQM0Vyy3XLEcajVLnnYRZa5fFb0xyWVXSCA"],
    # ["ブランディングエンジニア 成果", ""],
    # ["パフォテク 成果", ""],
    # ["アレテコ 成果", "17tndahR_B9XscWh_eE2orJPLZbGi32rILQoSLPD0oA0"],
    # ["ブリーチ 成果", ""],
    # ["FA 成果", "1SswfIyB-5tRaqEZlOXMzJ1L7PdFkAW-lqzBRelPdGvs"],
    # ["ナハト（クラン） 成果", "1IZiwYxSucjl-BAN3vV3F_P0-suJ1s0xXM2kcOrNWsQc"],
    # ["FORCE 成果", "1MHbIdekPMUJV5g8dHVyFyfmaJF3XggAxDd3Mvd2arzE"],
    # ["クリエイト 成果", ""],
    ["h1 成果", "1bTjfbr3qnStfRKcDvXuD-V8KE5z-5L0GKlF8RVwLUqw"],
    ["h1 LINE SEO", "142UZtsQ6N_Vp3m19x4oNLwgXeZW4pcDmeu5b9EJSe-I"],
    ["h3 成果", "1jhUvxyI7GuqRyn3l0BFkRN-zJSeuw9Tb6w2dG1Ahn1I"]
]


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

# 成果シートを更新するループの関数
def update_sheet(key_list_seika, name_to, df_origin):
    for this_key in key_list_seika:
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
# 基幹タブ更新
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

    for this_sheet in key_list_roas:
        sheet_name = this_sheet[0]
        sheet_key = this_sheet[1]

        # シートにアクセス
        wb = func.connect_gspread(sheet_key)

        # CSVをアップロード
        wb.values_update(
            '基幹_当月',
            params={'valueInputOption': 'USER_ENTERED'},
            body={'values': list(csv.reader(open(csv_filepass, encoding='cp932')))}
        )
        time.sleep(3)

        print(f'{sheet_name} 基幹更新完了')

    
    #########################
    # CSVファイルを削除
    #########################

    os.remove(csv_filepass)

    drive.quit()


else:
    print("月を跨いだので設定を更新してください")








########################
# 
# 成果確認シートの集客表タブ更新
# 
########################

# 集客表元データのシートにアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)

# 元データ取得
df_origin = pd.DataFrame(ws_origin.get_all_values()) # 元データ取得
df_origin = df_origin[df_origin[4] != ""]   # ID空白除去


# ループにする
update_sheet(key_list_seika, name_to, df_origin)

print("成果確認シート 完了")


# In[ ]:


########################
# 
# ROAS管理シートの集客表タブ更新
# 
########################

# シート名
name_to = "集客表_当月"


# 集客表元データのシートにアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)

# 元データ取得
df_origin = pd.DataFrame(ws_origin.get_all_values()) # 元データ取得
df_origin = df_origin[df_origin[4] != ""]   # ID空白除去


# ループにする
update_sheet(key_list_roas, name_to, df_origin)

print("ROAS確認シート 完了")

