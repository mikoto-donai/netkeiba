VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Race"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LAST_RACE_NUMBER_ As Integer = 12
Private url_parameters_() As String
Private race_dates_() As String
Private url_parameter_predictions_ As Object

Private Sub Class_Initialize()
    
End Sub

Private Function convertRacingDate(ByVal racing_date As String) As String
    
    Dim times As String
    times = Left(racing_date, InStr(racing_date, "回") - 1)
    
    Dim formatted_times As String
    formatted_times = Format(times, "00")

    Dim venue As String
    Dim event_date As String
    Select Case True
        Case racing_date Like "*札幌*"
            venue = "01"
            event_date = Replace(Right(racing_date, InStr(racing_date, "札幌")), "日目", "")
        Case racing_date Like "*函館*"
            venue = "02"
            event_date = Replace(Right(racing_date, InStr(racing_date, "函館")), "日目", "")
        Case racing_date Like "*福島*"
            venue = "03"
            event_date = Replace(Right(racing_date, InStr(racing_date, "福島")), "日目", "")
        Case racing_date Like "*新潟*"
            venue = "04"
            event_date = Replace(Right(racing_date, InStr(racing_date, "新潟")), "日目", "")
        Case racing_date Like "*東京*"
            venue = "05"
            event_date = Replace(Right(racing_date, InStr(racing_date, "東京")), "日目", "")
        Case racing_date Like "*中山*"
            venue = "06"
            event_date = Replace(Right(racing_date, InStr(racing_date, "中山")), "日目", "")
        Case racing_date Like "*中京*"
            venue = "07"
            event_date = Replace(Right(racing_date, InStr(racing_date, "中京")), "日目", "")
        Case racing_date Like "*京都*"
            venue = "08"
            event_date = Replace(Right(racing_date, InStr(racing_date, "京都")), "日目", "")
        Case racing_date Like "*阪神*"
            venue = "09"
            event_date = Replace(Right(racing_date, InStr(racing_date, "阪神")), "日目", "")
        Case racing_date Like "*小倉*"
            venue = "10"
            event_date = Replace(Right(racing_date, InStr(racing_date, "小倉")), "日目", "")
        Case Else
            Err.Raise number:=513, Description:="レース情報の変換でエラーが発生しました"
    End Select
    
    Dim formatted_event_date As String
    formatted_event_date = Format(event_date, "00")

    convertRacingDate = venue & formatted_times & formatted_event_date

End Function

Public Function createUrlParameters(ByVal elements As Variant)
    On Error GoTo ErrorHandler
    
    Dim racing_dates() As String
    racing_dates = Filter(elements, "日目")
    
    ReDim url_parameters_(LAST_RACE_NUMBER_)
    
    Dim i As Long: i = 1
    Dim j As Long: j = 1
    For i = 1 To UBound(racing_dates)
        For j = 1 To LAST_RACE_NUMBER_
            url_parameters_(i) = (Year(Now) & convertRacingDate(racing_dates(i)) & Format(j, "00"))
        Next
    Next

    Exit Function
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function

Public Function extractRaceDates(ByVal elements As Variant) As String()
    On Error GoTo ErrorHandler

    Dim i As Long: i = 1
    Dim j As Long: j = 1
    ReDim race_dates_(1)
    For i = 1 To UBound(elements)
        If elements(i) = "開催日程" Then
            Do
                race_dates_(UBound(race_dates_)) = elements(i + j)
                If InStr(elements(i + j + 1), "/") = 0 Then
                    Exit Do
                End If
                ReDim Preserve race_dates_(UBound(race_dates_) + 1)
                j = j + 1
            Loop
            Exit For
        End If
    Next

    extractRaceDates = race_dates_()
    
    Exit Function
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function

Private Function fetchRacePrediction(destination_sheet As Object)
'ThisWorkbook.Sheets(2).Range("$A$1")
    With ThisWorkbook.Sheets(2).QueryTables.Add(Connection:= _
        "URL;https://race.netkeiba.com/?pid=yoso&id=c" & url_parameter_prediction_, destination:=destination_sheet)
        .Name = "URL_NETKEIBA_PREDICTION_PARAMETER"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlAllTables
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With

End Function

Public Function fetchRacePredictions(url_parameter_predictions)

    url_parameter_predictions_ = url_parameter_predictions

    Dim i As Integer
    For i = 0 To race_dates_.counts - 1
        Debug.Print url_parameter_predictions_(i)
        fetchRacePrediction (url_parameter_predictions_(i))
    Next

End Function

