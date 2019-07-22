Attribute VB_Name = "Controller"
Option Explicit

Public Static Function main()
On Error GoTo ErrorHandler

    StaticModule.initialize
    
    Dim race_year As Long
    race_year = 2018
    
    Dim race_places As Object: Set race_places = CreateObject("Scripting.Dictionary")
    With race_places
        .Add "01", "D–y"
'        .Add "02", "”ŸŠÙ"
'        .Add "03", "•Ÿ“‡"
'        .Add "04", "VŠƒ"
'        .Add "05", "“Œ‹"
'        .Add "06", "’†R"
'        .Add "07", "’†‹"
'        .Add "08", "‹“s"
'        .Add "09", "ã_"
'        .Add "10", "¬‘q"
    End With
    
    Dim race_master_url As String
    Dim o_past_race As New PastRace
    Dim o_fetcher As New Fetcher
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

    StaticModule.finalize

    End
ErrorHandler:
    StaticModule.logContent Now & vbTab & "ì‹Æ‚ğ’†’f‚µ‚Ü‚µ‚½" & vbTab & Err.Number & ":" & Err.Description
    StaticModule.finalize
End Function
