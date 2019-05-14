# 概要

[netkeiba.com](https://race.netkeiba.com)から今週のレース予想を取得して、エクセルに出力します。

## 動作環境
下記環境にて動作確認済みです。
- Windows10 Professional
- Microsoft Excel 2010
- O365 SMTPサービス

## 手動実行
### Excel VBAエディタから直接実行  
1. /bin/netkeiba.xlsm を起動します。
1. 開発タブをからVBAエディタを起動します。
1. 標準モジュールの`controller.main()`メソッドを実行します。
1. 下記のようにディレクトリが作成されます。既にファイルが存在する場合は上書きされます。
1. 完了すると/logにログが出力されます。

出力フォルダ<sup>†1</sup>  
L__ Excelブック(ファイル名<sup>†2</sup>:直近のレース日、シート名:レース番号)  
L__ Excelブック(ファイル名:直近のレース日、シート名:レース番号)  
L__ ••• 
    
<sup>†1</sup>: 規定値はDesktop/{取得日}  
<sup>†2</sup>: 注目のレース日を起点とした、前後2日のレース日  

例  
Desktop/20190511  
L__ 1回新潟5日目.xlsx  
L__ 1回新潟6日目.xlsx  

### バックグラウンドプロセスで実行
`/src/tasks`をダブルクリックします。

## 自動実行
タスクスケジュールで`/src/tasks`を実行します。
金曜日、土曜日のそれぞれ23:00頃に実行すると、直近データの予想が揃います。

# 管理方法
## ソース管理
[vbac](https://github.com/vbaidiot/Ariawase)を利用して、テキストファイルの状態でGitで管理します。
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
