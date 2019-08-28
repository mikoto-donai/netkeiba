# 概要

[netkeiba.com](https://race.netkeiba.com)から指定年、指定場所のレース結果を取得して、エクセルに出力します。

## 動作環境
下記環境にて動作確認済みです。
- Windows10 Professional
- Microsoft Excel 2010

## 手動実行
### 事前準備
コマンドプロンプトを開きます。netkeibaのルートディレクトリで下記を実行して、ソースファイルから /bin/ に本体のExcelファイルを作成します。
 `cscript vbac.wsf combine`  
 
*本体のExcelの容量が多きくなった場合は破棄して、上記コマンドを実行してExcelファイルを再度作成します*  

### VBAエディタから実行  
1. /bin/netkeiba.xlsm を起動します。初回起動時は、マクロを有効化しておきます。
1. /user/user の1行目に、netkeiba有料アカウントのユーザーID、2行目にパスワードをそれぞれ入力します。  
(デフォルトはユーザーID:hoge, パスワード:fuga)  
*ユーザーID, パスワードが入力されていない場合は、タイム指数などのパラメーターが＊＊となります*

1. 開発タブからVBAエディタを起動して、標準モジュールのController.main()を表示します。
1. 取得対象年(race_year)を設定します。
1. 取得対象外の場所(race_places)をコメントアウトします。
1. 標準モジュールの`controller.main()`を実行します。
1. 1ファイル数十秒程度でファイルを作成します。ディレクトリ構成は下記です。既にファイルが存在する場合は上書きされます。  
完了すると'/log'にログが出力されます。

出力フォルダ<sup>†1</sup>  
L__ Excelブック(ファイル名:レース日、シート名:レース番号)  
L__ Excelブック(ファイル名:レース日、シート名:レース番号)  
L__ ••• 
    
<sup>†1</sup>規定値はDesktop/{対象年_場所}  

例  
Desktop/2018_札幌  
L__ 1回新潟5日目.xlsx  
L__ 1回新潟6日目.xlsx  

### バックグラウンドプロセスで実行
前述の 1 - 4 を設定後、
`/src/tasks`をダブルクリックします。

# 管理方法
## ソース管理
[vbac](https://github.com/vbaidiot/Ariawase)を利用して、テキストファイルの状態でGitで管理します。
- Excelからテキストファイルに `cscript vbac.wsf decombine`
- テキストファイルからExcelに `cscript vbac.wsf combine`


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
