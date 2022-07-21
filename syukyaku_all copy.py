#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 集客表のデータ取得


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
from datetime import datetime as dt

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
key_02 = "1WORdOLMmZU-7xyEhtZUJbjwyKe_q1Ye47E30xSA4ZS4"
# シート名
name_02 = "22年7月"

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
# 成果更新シートのまとめデータ更新
# 
########################

key_list = [
    ["成果更新シート", "1mrktRFIhx9ASA7nyVVsecAuUMNxRPKV0d54SNxDSImM"]
]

name_to = "まとめデータ"


# 集客表元データのシートにアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)

# 元データ取得
df_origin = pd.DataFrame(ws_origin.get_all_values()) # 元データ取得
df_origin = df_origin[df_origin[4] != ""]   # ID空白除去

# ループにする
update_sheet(key_list, name_to, df_origin)


print("成果更新シート 完了")


# In[3]:


########################
# 
# 成果確認シートの集客表タブ更新
# 
########################

key_list = [
    ["フォースリー 成果 7月", "1p-NEAiD3uHUp-Gcvzi6bm8LF9NJYDQ5Ju24HwOS8UYU"],
    ["リンクエッジ 成果 7月", "1XbV5cgwkc6VuXlCUOSviqE4IdkbpyDxqug88089POfE"],
    ["サルクルー 成果 7月", "1MBz8_qw20tFHbphiysOv0paqkVq9b5WL3GHVvbm3fDc"],
    ["FORCE 成果 7月", "1SqjCFNE-gn_5feLbXx3iqnsWSR4e8RQtsfj3mkOlYTk"],
    ["レントラ（FA） 成果 7月", "1DCYAWHVnZd3KtvP25QwdAi6gbOQA4FIN1kYpW6pRS6M"],
    ["セレス 成果 7月", "1WqTDLkMwHbSu5mnNZgJRjShaXqM6LC3Fi9ItPJdwxdI"],
    ["ブランディングエンジニア 成果 7月", "1RphlXXMpsnw9s7ovTzg96-qCBO7nF1-AwKh3xiHNix8"],
    ["パフォテク 成果 7月", "1KonsyJpx7MUJKdlt-d41xHMuVt2YbPRGJH_jzAgScsM"],
    ["アレテコ 成果 7月", "1d6mU8C7i3pFacOJrcgQfczZugHYe-B1pSjjYBDSWvyY"],
    ["ブリーチ 成果 7月", "1FsQAKFfPHTBsOM4s99eldTwb5LtnzQd2cibYEg_X1H4"],
    ["FA 成果 7月", "1L0DJRvcvu8x58nz0MWMB5Pu-JMrbk7D4RE3ep-ALioo"],
    ["クラン 成果 7月", "14GhIRCLUQWB-oPX3unXwyaqErJEOPEPwaHuhOAZwFIU"],
    ["クリエイト 成果 7月", "1BlStxYGDdMLrOZSfAqPzPpP4htbizGMEw6fy6_Ri6yw"],
    ["h1 成果 7月", "1-eU5PeNBNXqIONfIRZxJErxItERh6DhVlrx-9IewPhI"],
    ["h3 成果 7月", "1_JWsy811ICtYojOoVw_wm5lHLnpfXh0n8_r-zoIza3E"]
]

# シート名
name_to = "まとめデータ_22年7月"


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
    ["FORCE ROAS 7月", "1AXIhXCbKlzqFlS2cVaoiq0-H6bgIvtenB3lcZWsyTOA"],
    ["セレス ROAS 7月", "1Lpd8Orw8M2VlHAaEIuqqwGlwBtSSDnDa7tOo9o-zogM"],
    ["アレテコ ROAS 7月", "1rchXp8iuNTj6pyYCivK0_nC5a0V38GWXJNMJ4QhJxoE"],
    ["クラン ROAS 7月", "18s2UpmJ5z0EyMxepWVOZ7gDnQ8y5SOBF9Sma3P8Enws"]
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

print("成果確認シート 完了")

