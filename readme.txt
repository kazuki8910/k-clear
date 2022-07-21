◆ファイルの役割
・conf：変数格納
・func：関数格納
・kikan：基幹データ取り込み
・seika_quoriza：クオリザデータ取り込み
・syukyaku：集客表データ取り込み

◆ファイルの更新方法
・GitHubに更新
    git add .
    git commit -m "コメント"
    git push origin master
・herokuにアップ
    git push heroku master
    →エラーが起きたら heroku login でログインし直す
・herokuでの動作確認
    heroku run python syukyaku_all.py




◆Chromeのバージョン
Windows：97.0.4692.71
MAC：
heroku：98.0.4758.102


▼権限付与するメールアドレス
kazuki@store-system-k.iam.gserviceaccount.com


▼シート保存してるドライブ
https://drive.google.com/drive/u/0/folders/1rpxNAZqFJQQNnVpkSPzO90UyY8YDOYx-