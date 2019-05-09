VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RaceDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private current_race_dates_() As String
Private event_and_race_dates_() As String    '0420(“y)2‰ñ“Œ‹ž4“ú–Ú, 0420(“y)3‰ñ‹ž“s3“ú–Ú, 0421(“ú)2‰ñ“Œ‹ž5“ú–Ú, 0421(“ú)3‰ñ‹ž“s6“ú–Ú
Private Sub Class_Initialize()
    
End Sub
Public Property Get currentRaceDates() As String()
    currentRaceDates = current_race_dates_()
End Property

Public Function analyzeItems(ByRef current_race_event_parameters() As String)
On Error GoTo ErrorHandler
    
    If createCurrentRaceDates(current_race_event_parameters()) <> 0 Then
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "ƒGƒ‰["
    analyzeItems = 1
    
End Function

Private Function createCurrentRaceDates(ByRef current_race_event_parameters() As String) As Long
    
    ReDim current_race_dates_(0)
    Dim o_fetcher As New fetcher
    Dim i As Integer
    
    For i = LBound(current_race_event_parameters()) To UBound(current_race_event_parameters())
        o_fetcher.url = "https://race.netkeiba.com/?pid=race_list&id=" & current_race_event_parameters(i)
        o_fetcher.targetSheet = ThisWorkbook.Sheets(1)
        o_fetcher.fetchItems
        createCurrentRaceDate o_fetcher.items
    Next
    
    If UBound(current_race_dates_()) > 0 Then
        ReDim Preserve current_race_dates_(UBound(current_race_dates_()) - 1)
    End If
    
End Function

Private Function createCurrentRaceDate(ByRef items As Variant) As Long
    
    Dim i As Long
    For i = LBound(items, 1) To UBound(items, 1)
        If InStr(items(i, 1), "“ú–Ú") > 0 Then
            current_race_dates_(UBound(current_race_dates_())) = items(i, 1)
            ReDim Preserve current_race_dates_(UBound(current_race_dates_()) + 1)
        End If
    Next
    
    createCurrentRaceDate = 0
    
End Function

