Attribute VB_Name = "Controller"
Option Explicit

Public Static Function main()
On Error GoTo ErrorHandler
    
    StaticModule.initialize
    
    Dim race_year As Long: race_year = 2018
    Dim race_place As Long: race_place = 9
    
    Dim o_fetcher As New Fetcher
    Dim o_past_race As New PastRace

    o_fetcher.url = "https://keiba.yahoo.co.jp/schedule/list/" & CStr(race_year) & "/?" & "place=" & CStr(race_place)
    o_fetcher.fetchItems

    o_past_race.analyzeItems race_year, o_fetcher.items
 
    Dim o_directory As New Directory
    o_directory.fileNames = o_past_race.fileNames
    o_directory.contents = o_past_race.raceResults
    o_directory.createFiles
    
    StaticModule.finalize
    
    End
ErrorHandler:
    StaticModule.logContent Now & vbTab & "çÏã∆ÇíÜífÇµÇ‹ÇµÇΩ" & vbTab & Err.Number & ":" & Err.Description
    StaticModule.finalize
End Function

