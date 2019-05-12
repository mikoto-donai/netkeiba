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

Private current_race_dates_() As String     '注目レース開催日を起点とした直近のレース日　　ex. 2回東京4日目, 3回京都3日目

Public Property Get currentRaceDates() As String()
    currentRaceDates = current_race_dates_()
End Property

Public Function outputLog() As String
    outputLog = Now & vbTab & "対象レース日を作成しました: " & Join(current_race_dates_(), ",")
End Function

Public Function analyzeItems(ByRef current_race_event_parameters() As String)
    
    If createCurrentRaceDates(current_race_event_parameters()) = -1 Then
        Err.Raise Number:=513, Description:="カレントレース日の作成に失敗しました"
        Exit Function
    End If
    
End Function

Private Function createCurrentRaceDates(ByRef current_race_event_parameters() As String) As Long
    
    ReDim current_race_dates_(0)
    Dim o_fetcher As New Fetcher
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
    
    createCurrentRaceDates = 0
    
End Function

Private Function createCurrentRaceDate(ByRef items As Variant)
    
    Dim i As Long
    For i = LBound(items, 1) To UBound(items, 1)
        If InStr(items(i, 1), "日目") > 0 Then
            current_race_dates_(UBound(current_race_dates_())) = items(i, 1)
            ReDim Preserve current_race_dates_(UBound(current_race_dates_()) + 1)
        End If
    Next
    
End Function


