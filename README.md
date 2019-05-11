# 概要

## 動作環境(検証済)
- Windows10 Professional
- Microsoft Excel 2010
- O365 SMTPサービス

## 機能概要
### 動作

1. EXCELファイルの作成  
netkeiba.xlsm を起動してマクロが実行されると、下記のようにディレクトリが作成されます。  

  netkeiba.xlsm　　
  |  
  フォルダ(フォルダ名: 作成日)  
    |--- Excelブック(ファイル名:開催日、シート名:レース番号)  
    |--- Excelブック(ファイル名:開催日、シート名:レース番号)  
    |--- ••• 
    
例
  netkeiba.xlsm  
  |  
  20190511  
  |_ 20180501.xlsx (Sheet1:12)
  |_ 20180502.xlsx  


1. URL確認
下記サイトURLが存在するかを確認します  
http://race.netkeiba.com/?pid=yoso&id=c{year}{venue}{times}{event_date}{race_number}  

## バッチ処理
タスクスケジューラでバッチ処理を行います

金曜日、土曜日のそれぞれ23:00に、該当データが存在すればバッチ処理で下記の通り作成します。

- フォルダ名: 
- ファイル名: {year}{venue}{times}{event_date}
- シート名: {race}

### 結果の送信
- O365 SMTPサービスを利用します。  
  - アカウント: keiba.keiba@outlook.com  
  
- コマンドプロンプトでメール送信を行います。  
`powershell -NoProfile -ExecutionPolicy Unrestricted .\sendMailByO365.ps1`  


# コーディング規約
- 全体
  - プロシージャ(Sub)を使わない
  - メソッド(Function)で統一する
  - クラスに属しない関数は定義しない
  - 戻り値なし、実引数括弧ありのメソッドは強制値渡しになるので利用しない
  - 引数が2つ以上存在するメソッドは、名前付き引数を利用する
  - 仮引数は値渡し、参照渡しを明示する
  - mainルーチンをcontroller.mainに定義する
  - 参照設定を使わず、CeateObjectを利用する
  - レンダリング(セル操作)は一つのクラスにまとめる
  
- 命名規則
  - インスタンス変数にアンダースコア(_)をつける
  - クラス型変数の接頭辞にo_をつける
  - 定数は全て大文字とする
  - メソッド名はローワーキャメルとする
  - 変数名はスネークケースとする
  - 簡略した英単語はカウンタ変数のみとする
