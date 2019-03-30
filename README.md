## 概要

## 環境
- Windows10 Professional
- MS EXCEL
- O365 SMTPサービス

## 機能概要
### 動作

1. EXCELファイルの作成  
netkeiba.xlsm を起動してマクロが実行されると、下記のようにディレクトリが作成されます。  

  netkeiba.xlsm
  |  
  フォルダ(フォルダ名: システム日付_当日の累積起動回数)  
    |--- Excelブック(ファイル名:開催日、シート名:レース番号)  
    |--- Excelブック(ファイル名:開催日、シート名:レース番号)  
    |--- ••• 
    
例
  netkeiba.xlsm  
  |  
  20180901_01  
  |_ 20180501.xlsx (Sheet1:12)
  |_ 20180502.xlsx  

1. フォルダ作成  
Excelと同じ位置に、システム日付でフォルダを作成します  
フォルダ名{year}{month}{date}_{created_number}を作成します  

1. URL確認
下記サイトURLが存在するかを確認します  
http://race.netkeiba.com/?pid=yoso&id=c{year}{venue}{times}{date}{race}  

### URLパラメータ

データ取得対象のURLは下記です。  
http://race.netkeiba.com/?pid=yoso&id=c{year}{venue}{times}{date}{race}

URLの各変数は下記に従います。

| 変数名 | 説明 | 例 |
------|--------|-------| 
| year | 開催年を表します    |   2018    |
| venue | 開催場所を表します |  01:札幌 <br> 02:函館 <br> 03:福島 <br> 04:新潟 <br> 05:東京 <br> 06:中山 <br> 07:中京 <br> 08:京都 <br> 09:阪神 <br> 10:小倉 |
| times | 開催次数を表します | 01: 1回目  |
| date | 開催日を表します     | 02: 2日目  |
| race | 開催レースを表します  | 12: 12R   |

ex. 18/02/17 1回東京7日目
http://race.netkeiba.com/?pid=yoso&id=p201805010701 

## メール送信
- O365 SMTPサービスを利用します。  
  - アカウント: keiba.keiba@outlook.com  
  
- コマンドプロンプトでメール送信を行います。  
`powershell -NoProfile -ExecutionPolicy Unrestricted .\sendMailByO365.ps1`  

## バッチ処理
タスクスケジューラでバッチ処理を行います

金曜日、土曜日のそれぞれ23:00に、該当データが存在すればバッチ処理で下記の通り作成します。

- フォルダ名: {year}{month}{date}_{created_number}
- ファイル名: {year}{venue}{times}{date}
- シート名: {race}

* {create_number}
取得回数を表します。

