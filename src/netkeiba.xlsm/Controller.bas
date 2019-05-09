Attribute VB_Name = "Controller"
Option Explicit

Public Static Function main()

    Application.DisplayAlerts = False

    Dim o_fetcher As New fetcher
    o_fetcher.fetchItems

    Dim o_race_event As New raceEvent
    o_race_event.analyzeItems o_fetcher.items

    Dim o_race_date As New RaceDate
    o_race_date.analyzeItems o_race_event.current_race_event_parameters

    Dim o_prediction As New Prediction
    o_prediction.analyzeItems o_race_date.currentRaceDates
    
    Dim o_directory As New directory
    o_directory.fileNames = o_prediction.predictionRaceDates
    o_directory.contents = o_prediction.predictions
    o_directory.createFiles

    Application.DisplayAlerts = True

End Function
