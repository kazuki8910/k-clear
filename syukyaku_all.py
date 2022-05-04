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
from datetime import datetime as dt

# スプシ関連
from gspread_dataframe import set_with_dataframe

# 自作モジュール
import func
import conf

print("モジュールのインポート完了")


# In[3]:


########################
# 
# 成果確認シートの集客表タブ更新
# 
########################

# シートキー
key_02 = "1WORdOLMmZU-7xyEhtZUJbjwyKe_q1Ye47E30xSA4ZS4"

key_list = [
    ["フォースリー 成果 5月", "15Jn-q4G5060iDOII7EW8_zhDdWfwFHB3oWwQRMfanqs"],
    ["リンクエッジ 成果 5月", "1fsu-uILuZ7Y72GGX86dwmrJgXi20ZOQmWyftqIPRuco"],
    ["サルクルー 成果 5月", "1G5kxK7lmHHZJTP29N9ApRQV4TJoYT2RaVltMK_Wldew"],
    ["FORCE 成果 5月", "11nQNMGIzlfX8KM-jreCwnrjqdsAUfUliHgrwqdHIG4o"],
    ["レントラ（FA） 成果 5月", "1AjAHRleLKcHRWKnN1E77_NbBGsCmOIs3jLLAps4iwQE"],
    ["セレス 成果 5月", "1YwE-Hdx058OEkhHjovjal0ScKEJNWOIFEBtR-qoIyWA"],
    ["ブランディングエンジニア 成果 5月", "1huUp9ASQCPD5mnOYcrn5XcJzKVvCaQnUr3o7guNhcNU"],
    ["パフォテク 成果 5月", "1lENLjIfu49QANCQUNVmaB7A9Cd2Ha--nwnM4P2oh2yk"],
    ["アレテコ 成果 5月", "1h-eKxkFgKjbtU4zSZ27w_qX5xB2I43to0wyFmNXQ_x0"],
    ["ブリーチ 成果 5月", "1dlyj0kda6GpvNWLCpHPEU1h0kkMGSvE9798uVz_ibfk"],
    ["FA 成果 5月", "1f6DAv8sFpUio8ougLu4Sm_exrJV9D_0kF75dUbKqqEo"],
    ["クラン 成果 5月", "1yyNXq88-XQROgeQdt6LtEy5UretavethE1Q-mqgk_sQ"]
]

# シート名
name_02 = "22年5月"
name_to = "まとめデータ_22年5月"


# 集客表元データのシートにアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)

# 元データ取得
df_origin = pd.DataFrame(ws_origin.get_all_values()) # 元データ取得
df_origin = df_origin[df_origin[4] != ""]   # ID空白除去

# コピー先にアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)

# ループにする
for this_key in key_list:
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

print("成果確認シート 完了")


# In[ ]:


########################
# 
# 成果確認シートの集客表タブ更新
# 
########################

# シートキー
key_02 = "1WORdOLMmZU-7xyEhtZUJbjwyKe_q1Ye47E30xSA4ZS4"

key_list = [
    ["サルクルー ROAS 5月", "1c_OptSDdKnwPT2hCR5GeLx2zRN8XoBh0UIozH-dIuqw"],
    ["FORCE ROAS 5月", "1mkku6tfljMB8Eqhf-BO7FrREggMywUJXW3PakpgktTA"],
    ["セレス ROAS 5月", "1Cyco-l3AOUdMGU28KNI5lDhqLnyq4uzSWYY9_U7qRCE"],
    ["パフォテク ROAS 5月", "1ym1zIV-_ZFkNeCDS2JxvmOq8kbsQAvK-eTfyRYKSdVQ"],
    ["アレテコ ROAS 5月", "1-gN3Zgot1hiG8V1po4247jS1a8FzG6k3ZK6QL4P2ow4"],
    ["ブリーチ ROAS 5月", "1X6rNJoFEaRttEaihhbwRm57T9Ho2FOumtvQCAhWZVyQ"],
    ["FA ROAS 5月", "1h5o5m3lynyufizEpQTf-fXbR1WoeutwZHJUN6V9AITs"],
    ["クラン ROAS 5月", "1iKUNYXU9uwO35MplWyMh6Y4l7gEF78O7AQ2135WI7gk"],
    ["ナハト ROAS 5月", "1RRJDE6ZFrsAFl2UkQo5Tnrjg4QIXdSUXFgiSrgSN6FY"],
    ["エンジョイ ROAS 5月", "1UjXevYXR1vZTHHX1SvoTWyxP4mW2zvf0rLVAWjITm_E"],
    ["無限 ROAS 5月", "1tTF5NOdqN5CCLoUAbC9qxXVPEFH8c2UdVwqGGRRUiE4"]
]

# シート名
name_02 = "22年5月"
name_to = "集客表_当月"


# 集客表元データのシートにアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)

# 元データ取得
df_origin = pd.DataFrame(ws_origin.get_all_values()) # 元データ取得
df_origin = df_origin[df_origin[4] != ""]   # ID空白除去

# コピー先にアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)

# ループにする
for this_key in key_list:
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

print("成果確認シート 完了")

