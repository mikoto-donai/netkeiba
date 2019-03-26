## 概要

## 環境
- Windows10 Professional
- MS EXCEL
- O365 SMTPサービス

## 機能概要
### 動作

1. フォルダ、EXCELファイルの作成
本体Excelブックのマクロが実行されると、下記のようにディレクトリが作成されます。  
  本体Excelブック  
    |--- 日時フォルダ(フォルダ名: 日時_取得回数)  
    |--- Excelブック(ファイル名:開催日、シート:レース番号)  
    |--- Excelブック(ファイル名:開催日、シート:レース番号)  
    |--- •••  
    
1. Excelと同じ位置に、現在日でフォルダを作成します  
フォルダ名{year}{month}{date}_{created_number}を作成します  

1. 下記サイトURLが存在するかを確認します  
http://race.netkeiba.com/?pid=yoso&id=c{year}{place}{number}{date}{race}  

### URLパラメータ

データ取得対象のURLは下記です。  
http://race.netkeiba.com/?pid=yoso&id=c{year}{place}{number}{date}{race}

URLの各変数は下記に従います。

| 変数名 | 説明 | 例 |
------|--------|-------| 
| year | 開催年を表します    |   2018    |
| place | 開催場所を表します |  01:札幌 <br> 02:函館 <br> 03:福島 <br> 04:新潟 <br> 05:東京 <br> 06:中山 <br> 07:中京 <br> 08:京都 <br> 09:阪神 <br> 10:小倉 |
| number | 開催次数を表します | 01: 1回目  |
| date | 開催日を表します     | 02: 2日目  |
| race | 開催レースを表します  | 12: 12R   |

ex. 18/02/17 1回東京7日目
http://race.netkeiba.com/?pid=yoso&id=p201805010701 

## Excelファイルの作成
金曜日、土曜日のそれぞれ23:00に、該当データが存在すればバッチ処理で下記の通り作成します。

- フォルダ名: {year}{month}{date}_{created_number}
- ファイル名: {year}{place}{number}{date}
- シート名: {race}

* {create_number}
取得回数を表します。

例
フォルダ名: 20180501_1
ファイル名: 2018050101.xls
シート名: 1R, 2R, ・・・, 12R

## メール送信
- O365 SMTPサービスを利用します。  
  - アカウント: keiba.keiba@outlook.com  
  
- コマンドプロンプトでメール送信を行います。  
`powershell -NoProfile -ExecutionPolicy Unrestricted .\sendMailByO365.ps1`  

- タスクスケジューラでバッチ処理を行います

