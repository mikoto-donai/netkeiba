## 概要


## 仕様 

### 機能概要
#### フォルダ、EXCELファイルの作成

本体Excelブックのマクロが実行されると、下記のようにディレクトリが作成されます。

本体Excelブック  
    |--- 日時フォルダ(フォルダ名: 日時_取得回数)  
    |--- Excelブック(ファイル名:開催日、シート:レース番号)  
    |--- Excelブック(ファイル名:開催日、シート:レース番号)  
    |--- •••  

Excelと同じ位置に、現在日でフォルダを作成します  
フォルダ名{year}{month}{date}_{created_number}を作成します  

下記サイトURLが存在するかを確認します  
http://race.netkeiba.com/?pid=yoso&id=c{year}{place}{number}{date}{race}  

