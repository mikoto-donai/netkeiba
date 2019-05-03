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

Private race_dates_() As String      '1回新潟3日目, 2回東京5日目
Private prediction_parameters_() As String   'c201905020501, c201905020501
Private predictions_() As Variant    '?pid=yoso&id={prediction_parameter} で得られる予想データ / 2次元配列: predictions(race_date, race_number)

Public Property Let raceDates(ByRef race_dates() As String)
    race_dates_() = race_dates()
End Property

Public Property Get predictionParameters() As String()
    predictionParameters = prediction_parameters_()
End Property

Public Property Let predictions(ByRef predictions() As Variant)
    predictions_() = predictions()
End Property

Public Property Get predictions() As Variant()
    predictions = predictions_()
End Property

Public Function analyzeItems()
    
    If createPredictionParameters <> 0 Then
    End If
    
End Function

Private Function fetchRacePredictions(url_parameter_predictions)

    Dim i As Integer
    For i = 0 To race_dates_.counts - 1
        Debug.Print url_parameter_predictions_(i)
        fetchRacePrediction (url_parameter_predictions_(i))
    Next

End Function

Private Function createPredictionParameters() As Long

    Const LAST_RACE_NUMBER_ As Integer = 12
    ReDim prediction_parameters_(UBound(race_dates_()), LAST_RACE_NUMBER_)
    
    Dim i As Long
    Dim j As Long
    For i = LBound(race_dates_()) To UBound(race_dates_())
        For j = 0 To LAST_RACE_NUMBER_ - 1
            prediction_parameters_(i, j) = (Year(Now) & convertRaceDateIntoParameters(race_dates_(i)) & Format(j + 1, "00"))
        Next
    Next
   
    createPredictionParameters = 0

End Function

Private Function convertRaceDateIntoParameters(ByVal race_date As String) As String
    
    Dim times As String
    times = Left(race_date, InStr(race_date, "回") - 1)
    
    Dim formatted_times As String
    formatted_times = Format(times, "00")

    Dim venue As String
    Dim event_date As String
    Select Case True
        Case race_date Like "*札幌*"
            venue = "01"
            event_date = Replace(Right(race_date, InStr(race_date, "札幌")), "日目", "")
        Case race_date Like "*函館*"
            venue = "02"
            event_date = Replace(Right(race_date, InStr(race_date, "函館")), "日目", "")
        Case race_date Like "*福島*"
            venue = "03"
            event_date = Replace(Right(race_date, InStr(race_date, "福島")), "日目", "")
        Case race_date Like "*新潟*"
            venue = "04"
            event_date = Replace(Right(race_date, InStr(race_date, "新潟")), "日目", "")
        Case race_date Like "*東京*"
            venue = "05"
            event_date = Replace(Right(race_date, InStr(race_date, "東京")), "日目", "")
        Case race_date Like "*中山*"
            venue = "06"
            event_date = Replace(Right(race_date, InStr(race_date, "中山")), "日目", "")
        Case race_date Like "*中京*"
            venue = "07"
            event_date = Replace(Right(race_date, InStr(race_date, "中京")), "日目", "")
        Case race_date Like "*京都*"
            venue = "08"
            event_date = Replace(Right(race_date, InStr(race_date, "京都")), "日目", "")
        Case race_date Like "*阪神*"
            venue = "09"
            event_date = Replace(Right(race_date, InStr(race_date, "阪神")), "日目", "")
        Case race_date Like "*小倉*"
            venue = "10"
            event_date = Replace(Right(race_date, InStr(race_date, "小倉")), "日目", "")
        Case Else
            Err.Raise number:=513, Description:="レース情報の変換でエラーが発生しました"
    End Select
    
    Dim formatted_event_date As String
    formatted_event_date = Format(event_date, "00")

    convertRaceDateIntoParameters = venue & formatted_times & formatted_event_date

End Function





