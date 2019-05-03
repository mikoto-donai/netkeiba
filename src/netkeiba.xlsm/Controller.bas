Attribute VB_Name = "Controller"
Option Explicit

Public Static Function main()

    Dim f_event As New Fetcher
    f_event.targetSheet = ThisWorkbook.Sheets(1)
    f_event.url = "https://race.netkeiba.com/?pid=top"
    f_event.fetchItems

    Dim e_event As New EventDate
    e_event.items() = f_event.items()
    e_event.analyzeItems
 
    Dim d_directory As New directory
    
    d_directory.outputPath = Environ("HOMEPATH") & "\" & "Desktop"
    d_directory.folderName = Format(Date, "yyyymmdd")

    Dim r_race As New RaceDate
    Dim f_race As New Fetcher
    
    Dim i As Long
    For i = LBound(e_event.current_event_dates()) To UBound(e_event.current_event_dates())
        
        f_race.url = "https://race.netkeiba.com/?pid=race_list&id=" & e_event.current_event_dates_parameters()(i)
        f_race.targetSheet = ThisWorkbook.Sheets(1)
        f_race.fetchItems
        r_race.items() = f_race.items()
        r_race.EventDate = e_event.current_event_dates()(i)
        r_race.analyzeItems
        
        
        d_directory.fileNames() = r_race.eventAndRaceDates()
        d_directory.createFiles
        
    Next


End Function


'Function test2()
'    Dim p_prediction As New Prediction
'    Dim race_dates() As String
'    ReDim race_dates(1)
'    race_dates(0) = "2‰ñ“Œ‹ž4“ú–Ú"
'    race_dates(1) = "2‰ñ“Œ‹ž5“ú–Ú"
'    p_prediction.raceDates() = race_dates()
'    p_prediction.analyzeItems
'    Debug.Print p_prediction.predictionParameters()(0, 3)
'
'    Dim buf() As Variant
'    ReDim buf(1, 11)
'    buf(0, 1) = Range("A1:A20").Value
'    buf(1, 2) = Range("A5:A20").Value
'    p_prediction.predictions = buf
'    Range("F6:F10").Value = p_prediction.predictions()(1, 2)
'
'End Function
