Attribute VB_Name = "Controller"
Option Explicit

Public Static Function main()
On Error GoTo ErrorHandler

    Dim o_configuration As New Configuration

    Dim o_fetcher As New Fetcher
    o_fetcher.fetchItems

    Dim o_race_event As New RaceEvent
    o_race_event.analyzeItems o_fetcher.items

    Dim o_race_date As New RaceDate
    o_race_date.analyzeItems o_race_event.current_race_event_parameters
    o_configuration.logContet = o_race_date.outputLog

    Dim o_prediction As New Prediction
    o_prediction.analyzeItems o_race_date.currentRaceDates
    
    Dim o_directory As New Directory
    o_directory.sheetNames = o_prediction.predictionRaceDates
    o_directory.contents = o_prediction.predictions
    o_directory.createFiles
    
    o_configuration.finalize
    
    End
ErrorHandler:
    o_configuration.logContet = Now & vbTab & "çÏã∆ÇíÜífÇµÇ‹ÇµÇΩ" & vbTab & Err.Number & ":" & Err.Description
    o_configuration.finalize
End Function

