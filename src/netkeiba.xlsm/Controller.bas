Attribute VB_Name = "Controller"
Option Explicit

Public Static Function main()

    Dim o_view As New View
    Dim fetched_item_race_lists() As String
    fetched_item_race_lists() = o_view.fetchItems("https://race.netkeiba.com/?pid=top")

    Dim o_race As New race
    Dim current_race_list_parameters()  As String
    current_race_list_parameters() = o_race.createCurrentRaceListParameters(fetched_item_race_lists())
    
    Debug.Print current_race_list_parameters(LBound(current_race_list_parameters()))
    Debug.Print current_race_list_parameters(UBound(current_race_list_parameters()))
    
'    Dim race_dates() As String
'    race_dates() = o_race.createRaceDates(fetched_item_race_lists())
    
'    Dim fetched_item_current_races()
'    fetched_items_current_races() = o_view.fetchItems("https://race.netkeiba.com/?pid=")
'
'
'    Dim output_path As String
'    output_path = Environ("HOMEPATH") & "\" & "Desktop"
'
'    Dim o_directory As New directory
'    o_directory.cofigureOutputPath (output_path)
'
'    Dim created_date As String
'    created_date = Format(Date, "yyyymmdd")
'
    
'
'    If o_directory.createFiles(folder_name:=created_date, file_names:=race_dates()) = 1 Then
'        Debug.Print ""
'    End If
    
    'https://race.netkeiba.com/?pid=yoso&id=p201905020101
    
End Function
