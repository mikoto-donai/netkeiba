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

Private items_() As String
Private event_date_ As String    '5/04(“ú)
Private race_dates_() As String      '2‰ñ“Œ‹4“ú–Ú, 3‰ñ‹“s3“ú–Ú,
Private event_and_race_dates_() As String    '0420(“y)2‰ñ“Œ‹4“ú–Ú, 0420(“y)3‰ñ‹“s3“ú–Ú, 0421(“ú)2‰ñ“Œ‹5“ú–Ú, 0421(“ú)3‰ñ‹“s6“ú–Ú

Private Sub Class_Initialize()
    
End Sub

Public Property Let items(ByRef items() As String)
    items_() = items()
End Property

Public Property Let EventDate(ByVal event_date As String)
    event_date_ = event_date
End Property

Public Property Get eventAndRaceDates() As String()
    eventAndRaceDates = event_and_race_dates_()
End Property

Public Function analyzeItems()
On Error GoTo ErrorHandler
    
    If createRaceDates() <> 0 Then
    End If
    
    If createEventAndRaceDates() <> 0 Then
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "ƒGƒ‰["
    analyzeItems = 1
    
End Function

Private Function createRaceDates() As Long
    
    race_dates_() = Filter(items_(), "“ú–Ú")
    createRaceDates = 0
    
End Function

Private Function createEventAndRaceDates() As Long
    
    ReDim event_and_race_dates_(UBound(race_dates_()))
    
    Dim i As Long
    For i = LBound(race_dates_()) To UBound(race_dates_())
        event_and_race_dates_(i) = Replace(event_date_, "/", "") & race_dates_(i)
    Next
    
    createEventAndRaceDates = 0

End Function

