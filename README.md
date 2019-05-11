# 概要

netkeiba.com](https://race.netkeiba.com)から今週のレース予想を取得して、エクセルに出力します。

## 動作環境(検証済)
- Windows10 Professional
- Microsoft Excel 2010
- O365 SMTPサービス

## 動作方法(手動)
  
1. /bin/netkeiba.xlsm を起動します。
1. 標準モジュールの`controller.main()`メソッドを実行します。
1. 下記のようにディレクトリが作成されます。  

†1出力フォルダ  
L__ Excelブック(ファイル名:†2直近のレース日、シート名:レース番号)  
L__ Excelブック(ファイル名:直近のレース日、シート名:レース番号)  
L__ ••• 
    
†1: 規定値はDesktop/{取得日}  
†2: 注目のレース日を起点とした、前後2日のレース日  

例  
Desktop/20190511  
L__ 1回新潟5日目.xlsx  
L__ 1回新潟6日目.xlsx  

## 動作方法(自動)
タスクスケジューラでバッチ処理を行います

金曜日、土曜日のそれぞれ23:00に、該当データが存在すればバッチ処理で自動実行します。  
結果はメール送信します。

### メール送信
- O365 SMTPサービスを利用します。  
  - アカウント: keiba.keiba@outlook.com  
  
- コマンドプロンプトでメール送信を行います。  
`powershell -NoProfile -ExecutionPolicy Unrestricted .\sendMailByO365.ps1`  


# 管理方法
## ソース管理
[vbac](https://github.com/vbaidiot/Ariawase)を利用して、テキストファイルの状態でGitで管理する
- Excelからテキストファイルに `cscript /vbac.wsf decombine`
- テキストファイルからExcelに `cscript /vbac.wsf combine`


## コーディング規約
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
