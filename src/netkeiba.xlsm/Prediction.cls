VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Prediction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum RACE_NUMBER    'レース番号
    FIRST = 1
    LAST = 12
End Enum

Private prediction_parameters_() As Variant   '各レースの競馬予想URLパラメータ_(race_date, race_number)   ex.  prediction_parameters_(0, 10) / 201905020701
Private prediction_race_dates_() As String    'pid=yoso&id={prediction_parameter} で得られる競馬予想データ_(race_date, race_number)    ex.  predictions_race_dates(0, 10) / 2回東京7日目:11R
Private predictions_() As Variant    'pid=yoso&id={prediction_parameter} で得られる競馬予想データ_(race_date, race_number)    ex.  predictions_(0, 10) / 2回東京7日目:11R の予想データRangeオブジェクト

Public Property Get predictionRaceDates() As String()
    predictionRaceDates = prediction_race_dates_()
End Property

Public Property Get predictions() As Variant()
    predictions = predictions_()
End Property

Public Function analyzeItems(ByRef current_race_dates() As String)
    
    If createPredictionParameters(current_race_dates()) = -1 Then
        Err.Raise Number:=513, Description:="各レースの競馬予想URLパラメータの作成に失敗しました"
        Exit Function
    End If
    
    If fetchPredictions(current_race_dates()) = -1 Then
        Err.Raise Number:=513, Description:="各レースの競馬予想データの作成に失敗’しました"
        Exit Function
    End If

End Function

Private Function createPredictionParameters(ByRef current_race_dates() As String) As Long

    ReDim prediction_parameters_(UBound(current_race_dates()), RACE_NUMBER.LAST - 1)
    
    Dim i As Long
    Dim j As Long
    For i = LBound(current_race_dates()) To UBound(current_race_dates())
        For j = 0 To RACE_NUMBER.LAST - 1
            prediction_parameters_(i, j) = (Year(Now) & convertRaceDateIntoParameters(current_race_dates(i)) & Format(j + 1, "00"))
        Next
    Next
   
    createPredictionParameters = 0

End Function

Private Function convertRaceDateIntoParameters(ByVal current_race_date As String) As String
    
    Dim times As String
    times = Left(current_race_date, InStr(current_race_date, "回") - 1)
    
    Dim formatted_times As String
    formatted_times = Format(times, "00")

    Dim venue As String
    Dim event_date As String
    Select Case True
        Case current_race_date Like "*札幌*"
            venue = "01"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "札幌")), "日目", "")
        Case current_race_date Like "*函館*"
            venue = "02"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "函館")), "日目", "")
        Case current_race_date Like "*福島*"
            venue = "03"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "福島")), "日目", "")
        Case current_race_date Like "*新潟*"
            venue = "04"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "新潟")), "日目", "")
        Case current_race_date Like "*東京*"
            venue = "05"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "東京")), "日目", "")
        Case current_race_date Like "*中山*"
            venue = "06"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "中山")), "日目", "")
        Case current_race_date Like "*中京*"
            venue = "07"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "中京")), "日目", "")
        Case current_race_date Like "*京都*"
            venue = "08"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "京都")), "日目", "")
        Case current_race_date Like "*阪神*"
            venue = "09"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "阪神")), "日目", "")
        Case current_race_date Like "*小倉*"
            venue = "10"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "小倉")), "日目", "")
        Case Else
            Err.Raise Number:=513, Description:="レース情報の変換でエラーが発生しました"
    End Select
    
    Dim formatted_event_date As String
    formatted_event_date = Format(event_date, "00")

    convertRaceDateIntoParameters = venue & formatted_times & formatted_event_date

End Function

Private Function fetchPredictions(ByRef current_race_dates() As String) As Long

    ReDim prediction_race_dates_(UBound(prediction_parameters_, 1), UBound(prediction_parameters_, 2))
    ReDim predictions_(UBound(prediction_parameters_, 1), UBound(prediction_parameters_, 2))

    Dim i As Integer
    Dim j As Integer
    Dim o_fetcher As New Fetcher
    Const web_selection_type As String = xlAllTables
    
    For i = LBound(prediction_parameters_, 1) To UBound(prediction_parameters_, 1)
        For j = LBound(prediction_parameters_, 2) To UBound(prediction_parameters_, 2)
            
            o_fetcher.url = "https://race.netkeiba.com/?pid=yoso&id=c" & prediction_parameters_(i, j)
            o_fetcher.targetSheet = ThisWorkbook.Sheets("fetched_data")
            o_fetcher.fetchItems web_selection_type
            
            If Not IsEmpty(o_fetcher.items) Then
                predictions_(i, j) = o_fetcher.items
            End If
            
            prediction_race_dates_(i, j) = current_race_dates(i) & ":" & j + 1 & "R"
            
        Next
    Next
    
    fetchPredictions = 0
    
End Function




