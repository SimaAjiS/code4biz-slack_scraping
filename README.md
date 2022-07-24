# code4biz-slack_scraping
- code4biz Slackの内容をスクレイピングしエクセル保存するコード
- Python 3.9(pygraphbiz 1.7が3.10未対応)
- 構成
  - src　スクレイピングした日毎、月毎、期間毎のエクセルファイルを保存する
    - 集計.xlsx　最終アウトプットファイル
    - auth.json　code4biz Slack Webページへのログイン情報
  - scraping.py　指定した期間でスクレイピング
  - summary.py　日毎のエクセルファイルを一つにまとめる