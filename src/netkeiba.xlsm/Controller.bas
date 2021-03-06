Attribute VB_Name = "Controller"
Option Explicit

Public Static Function main()
On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    FileIO.startLog
    
    Dim race_year As Long
    race_year = 2018  '取得対象年を設定して下さい

    Dim race_places As Object: Set race_places = CreateObject("Scripting.Dictionary")
    With race_places  '取得対象外の開催場所をコメントアウトしてください
'        .Add "01", "札幌"
'        .Add "02", "函館"
'        .Add "03", "福島"
'        .Add "04", "新潟"
        .Add "05", "東京"
'        .Add "06", "中山"
'        .Add "07", "中京"
'        .Add "08", "京都"
'        .Add "09", "阪神"
'        .Add "10", "小倉"
    End With
    
    Const L1 = 1: Const L2 = 2
    Dim o_netkeiba As New Netkeiba
    o_netkeiba.login_id = FileIO.readFile(L1)
    o_netkeiba.login_pass = FileIO.readFile(L2)

    Const targetSheetName = "fetched_data"
    Dim o_fetcher As New Fetcher
    o_fetcher.configureTargetSheet targetSheetName

    Dim race_master_url As String
    Dim o_directory As New Directory

    Dim key As Variant
    For Each key In race_places
        race_master_url = "https://keiba.yahoo.co.jp/schedule/list/" & CStr(race_year) & "/?" & "place=" & key
        o_fetcher.url = race_master_url
        o_fetcher.fetchItems

        o_netkeiba.analyzeItems race_year, o_fetcher.items

        o_directory.folderName = race_year & "_" & race_places.Item(key)
        o_directory.fileNames = o_netkeiba.fileNames
        o_directory.contents = o_netkeiba.raceResults
        o_directory.createFiles
    Next

    FileIO.stopLog
    Application.DisplayAlerts = True

    End
ErrorHandler:
    FileIO.writeLog "作業を中断しました" & vbTab & Err.Number & ":" & Err.Description
    FileIO.stopLog
    Application.DisplayAlerts = True
End Function
