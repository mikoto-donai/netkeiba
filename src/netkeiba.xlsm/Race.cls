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

Private Const URL_NETKEIBA_RACE_LIST_ As String = "https://race.netkeiba.com/?pid=race_list"
Private Const LAST_RACE_NUMBER_ As Integer = 12
Private url_parameters_ As Object
Private race_dates_ As Object

Private Sub Class_Initialize()
    Sheets("Sheet1").Cells.Clear
    fetchRaceList
End Sub

Private Function fetchRaceList()

    With ThisWorkbook.Sheets(1).QueryTables.Add(Connection:= _
        "URL;" & URL_NETKEIBA_RACE_LIST_, Destination:=Range( _
        "$A$1"))
        .Name = URL_NETKEIBA_RACE_LIST_
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = True
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
End Function

Private Function convertRacingDate(racing_date As String) As String
    
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

Private Function makeBaseURLParameters() As Object

    Dim racing_date As Range
    Set racing_date = Cells.Find(What:="*回*日目")
   
    If racing_date Is Nothing Then
         Err.Raise number:=513, Description:="レース情報を取得できませんでした"
    End If
    
    Dim next_racing_date As Range: Set next_racing_date = racing_date
    Dim base_url_parameters_ As Object
    Set base_url_parameters_ = CreateObject("System.Collections.ArrayList")
    
    base_url_parameters_.Add (convertRacingDate(racing_date.Value))
    
    Do
        Set next_racing_date = Cells.FindNext(After:=next_racing_date)
        
        If next_racing_date Is Nothing Then
            Exit Do
        End If
 
        If next_racing_date.Address = racing_date.Address Then
            Exit Do
        End If
        
        base_url_parameters_.Add (convertRacingDate(next_racing_date.Value))
    Loop

    Set makeBaseURLParameters = base_url_parameters_

End Function

Private Function makeURLParameters(base_url_parameters As Object) As Integer

    Set url_parameters_ = CreateObject("System.Collections.ArrayList")
    
    Dim i As Integer: i = 0
    Dim j As Integer: j = 1
    For i = 0 To base_url_parameters.Count - 1
        For j = 1 To LAST_RACE_NUMBER_
            url_parameters_.Add (Year(Now) & base_url_parameters(i) & Format(j, "00"))
        Next
    Next
    
    makeURLParameters = url_parameters_.Count

End Function

Public Function defineURLs()
    On Error GoTo ErrorHandler
    
    If makeURLParameters(makeBaseURLParameters) = 0 Then
        Err.Raise number:=513, Description:="URLを作成できませんでした"
    End If

    Exit Function
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function

Public Function defineRaceDates() As Object
    On Error GoTo ErrorHandler

    Dim title_race_date As Range
    Set title_race_date = Cells.Find(What:="開催日程")

    If title_race_date Is Nothing Then
        Err.Raise number:=513, Description:="開催日程がありません"
    End If
    
    Set race_dates_ = CreateObject("System.Collections.ArrayList")
    Dim nextRow As Integer: nextRow = 1
    Do
        If InStr(Sheets(1).Cells(title_race_date.Cells.Row + nextRow, 1).Value, "/") = 0 Then
            Exit Do
        Else
            race_dates_.Add (Sheets(1).Cells(title_race_date.Cells.Row + nextRow, 1).Value)
            nextRow = nextRow + 1
        End If
    Loop

    Set defineRaceDates = race_dates_
    Exit Function
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function

Public Function fetchRacePrediction(URL_NETKEIBA_PREDICTION_PARAMETER As String)

    With ThisWorkbook.Sheets(2).QueryTables.Add(Connection:= _
        "URL;https://race.netkeiba.com/?pid=yoso&id=c" & URL_NETKEIBA_PREDICTION_PARAMETER, Destination:=ThisWorkbook.Sheets(2).Range("$A$1"))
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

