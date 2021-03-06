﻿# report_mgmt
An MS Excel addin for maintaining regular reports.

<インストール方法>

1. report_mgmt_tool.xlsaをダウンロード

2. レポートファイルを開く(レポートファイルのサンプルはsample.xls)

3. [ファイル]->[オプション]->[リボンのユーザ設定]で開発タブを有効化

4. [開発]->[アドイン]->[参照]でreport_mgmt_tool.xlsaを追加

5. [ファイル]->[オプション]->[クイックアクセスツールバー]で、
   [コマンドの選択]プルダウンメニューで「マクロ」を選択し、以下のマクロを追加

   - CreateRegularReport
   - AdjustTextWidth

   ( [ファイル]->[オプション]->[リボンのユーザ設定]でリボンへのマクロ追加も可能)


<使用方法>

1. レポートファイルのシートフォーマット

|       |  A列            |  B列        | C列 ...        |
| ----- | --------------- | ----------- | -------------- |
| 1行目 | 報告メール題目　| 項目1題目　 | 項目2題目  ... |
| 2行目 | 日付1           | 項目1内容1  | 項目1内容1 ... |
| 3行目 | 日付2           | 項目1内容2  | 項目1内容2 ... |
|　:    | :               |     :       |      :         |

- A列の表示形式は「日付」でなければならない。
- 本日の日付の入力は、TODAY()関数で自動取得&コピーした後、値だけペーストするとよい。
  (TODAY()関数のままだと、常にシートを開いた日の日付に更新されてしまう)
- セルに手動で改行を入れるにはAlt-ENTERを使用する。


2. セル内文字列整形

AdjustTextWidthマクロを用いて、セルの各行が最大文字数以内になるように改行を入れる。
- １行の最大文字数は、環境変数TEXT_WIDTH_IN_BYTESで指定可能
- １行の最大文字数のデフォルト値は80

3. Outlookメール下書き生成

CreateRegularReportマクロを用いて、レポートファイルからOutlookメール下書きを生成する。
  
- レポートメールのテンプレートは、環境変数REPORT_MAIL_TEMPLATE_FILEで指定可能。
- テンプレートの指定がない場合、宛先が空のレポートメールを生成。
- メールの題目はA列1行目のセルから自動生成される。
- メールの本文は日付が今日にマッチする行から自動生成される。
