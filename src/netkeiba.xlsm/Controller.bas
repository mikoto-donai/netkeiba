Attribute VB_Name = "Controller"
Option Explicit

Public Static Function main()
On Error GoTo ErrorHandler
    
    Logger.initialize
    Application.DisplayAlerts = False
    
    Dim race_year As Long
    race_year = 2018  '�擾�Ώ۔N��ݒ肵�ĉ�����
    
    Dim race_places As Object: Set race_places = CreateObject("Scripting.Dictionary")
    With race_places  '�擾�ΏۊO�̊J�Ïꏊ���R�����g�A�E�g���Ă�������
'        .Add "01", "�D�y"
'        .Add "02", "����"
'        .Add "03", "����"
'        .Add "04", "�V��"
        .Add "05", "����"
'        .Add "06", "���R"
'        .Add "07", "����"
'        .Add "08", "���s"
'        .Add "09", "��_"
'        .Add "10", "���q"
    End With
    
    Const targetSheetName = "fetched_data"
    Dim o_fetcher As New Fetcher
    o_fetcher.configureTargetSheet targetSheetName
    
    Dim race_master_url As String
    Dim o_past_race As New PastRace
    Dim o_directory As New Directory
    
    Dim key As Variant
    For Each key In race_places
        race_master_url = "https://keiba.yahoo.co.jp/schedule/list/" & CStr(race_year) & "/?" & "place=" & key
        o_fetcher.url = race_master_url
        o_fetcher.fetchItems
        
        o_past_race.analyzeItems race_year, o_fetcher.items
        
        o_directory.folderName = race_year & "_" & race_places.Item(key)
        o_directory.fileNames = o_past_race.fileNames
        o_directory.contents = o_past_race.raceResults
        o_directory.createFiles
    Next

    Logger.finalize
    Application.DisplayAlerts = True

    End
ErrorHandler:
    Logger.writeLog "��Ƃ𒆒f���܂���" & vbTab & Err.Number & ":" & Err.Description
    Logger.finalize
    Application.DisplayAlerts = True
End Function
