Attribute VB_Name = "Controller"
Option Explicit

Public Static Function main()

    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(1).Activate
    
    Dim o_view As New View
    Dim fetched_items() As String
    fetched_items() = o_view.fetchItems("https://race.netkeiba.com/?pid=race_list")

    Dim o_race As New race
    o_race.createUrlParameters (fetched_items())

'    o_race.fetchRacePrediction ("201906030608")
    
    Dim output_path As String
    output_path = Environ("HOMEPATH") & "\" & "Desktop"

    Dim o_directory As New directory
    o_directory.cofigureOutputPath (output_path)

    Dim created_date As String
    created_date = Format(Date, "yyyymmdd")
    
    Dim race_dates() As String
    race_dates() = o_race.extractRaceDates(fetched_items())

    If Not o_directory.createFiles(folder_name:=created_date, file_names:=race_dates) Then
'        o_race.fetchRacePredictions
    End If

'    bh.setBookNames = bookNames
'    Call bh.createFoldersAndBooks
    
    Application.DisplayAlerts = True
    Exit Function
    
End Function
