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
name_02 = "22年6月"

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
    ["フォースリー 成果 6月", "15Jn-q4G5060iDOII7EW8_zhDdWfwFHB3oWwQRMfanqs"],
    ["リンクエッジ 成果 6月", "1fsu-uILuZ7Y72GGX86dwmrJgXi20ZOQmWyftqIPRuco"],
    ["サルクルー 成果 6月", "1G5kxK7lmHHZJTP29N9ApRQV4TJoYT2RaVltMK_Wldew"],
    ["FORCE 成果 6月", "11nQNMGIzlfX8KM-jreCwnrjqdsAUfUliHgrwqdHIG4o"],
    ["レントラ（FA） 成果 6月", "1AjAHRleLKcHRWKnN1E77_NbBGsCmOIs3jLLAps4iwQE"],
    ["セレス 成果 6月", "1YwE-Hdx058OEkhHjovjal0ScKEJNWOIFEBtR-qoIyWA"],
    ["ブランディングエンジニア 成果 6月", "1huUp9ASQCPD5mnOYcrn5XcJzKVvCaQnUr3o7guNhcNU"],
    ["パフォテク 成果 6月", "1lENLjIfu49QANCQUNVmaB7A9Cd2Ha--nwnM4P2oh2yk"],
    ["アレテコ 成果 6月", "1h-eKxkFgKjbtU4zSZ27w_qX5xB2I43to0wyFmNXQ_x0"],
    ["ブリーチ 成果 6月", "1dlyj0kda6GpvNWLCpHPEU1h0kkMGSvE9798uVz_ibfk"],
    ["FA 成果 6月", "1f6DAv8sFpUio8ougLu4Sm_exrJV9D_0kF75dUbKqqEo"],
    ["クラン 成果 6月", "1yyNXq88-XQROgeQdt6LtEy5UretavethE1Q-mqgk_sQ"],
    ["クリエイト 成果 6月", "19aeLHkG6XxNKRQZD1yFnlRMtesdVWhjvKcEEB2xQJAk"],
    ["h1 成果 6月", "1pGgPvqZc2WSnUw0ZCcQHGvI8xciumtVopoYYaYzUhFY"],
    ["h3 成果 6月", "1yNs3crJD4-HlSG7bhpv0SQGlXncPZPOxmpb3jR4BSCg"]
]

# シート名
name_to = "まとめデータ_22年6月"


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
    ["FORCE ROAS 6月", "1ksJdBJ4q7H6UyjaO_2EvnxIqv4wGzIVKfDnfLEdxhDA"],
    ["セレス ROAS 6月", "1CeNbPf0cyF8baBknIuEeGUqEv8Namt4xUuOmhkyB1zQ"],
    ["パフォテク ROAS 6月", "1YMTBcjjMlhJCm7e3yD4l_opsiDzVbBz66ChyWRy1lTQ"],
    ["アレテコ ROAS 6月", "1R8W8fhSsjKgxh_TKsk-Vwsk2Mff1Gg6bADCb7gVzW5s"],
    ["ナハト ROAS 6月", "1RnlvVXbeIJCoo_HoMMAaQtyixdZ7-afmMTGEXqUtO4M"],
    ["エンジョイ ROAS 6月", "1bIcvd2H07OvOTgpO4wDZ0rwXwbVq2f3hEXp_Ig0bNTY"],
    ["無限 ROAS 6月", "1hGU-MxujS_hDN0pcrQL5iYQdhAITO72_sYl5sRCSp-0"],
    ["ウィンクレス ROAS 6月", "14t5-pW-2lJ_nEW4mlCagct-B8MxjZm_lFxmMMEF67BM"],
    ["クラン ROAS 6月", "1PkcqfaWlZufXuS6j1iyfBj827zNfQ4KopqjQclTP75E"]
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

