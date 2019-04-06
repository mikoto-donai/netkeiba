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

Private Const url_netkeiba_race_list As String = "https://race.netkeiba.com/?pid=race_list"
Private Const LAST_RACE_NUMBER As Integer = 12
Private url_parameters As Object
Private race_dates As Object

Private Sub Class_Initialize()
    Sheets("Sheet1").Cells.Clear
    fetchRaceList
End Sub

Private Function fetchRaceList()

    With ThisWorkbook.Sheets(1).QueryTables.Add(Connection:= _
        "URL;" & url_netkeiba_race_list, Destination:=Range( _
        "$A$1"))
        .Name = url_netkeiba_race_list
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

Private Function convertRacingDate(racing_date_ As String) As String
    
    Dim times As String
    times = Left(racing_date_, InStr(racing_date_, "回") - 1)
    
    Dim formatted_times As String
    formatted_times = Format(times, "00")

    Dim venue As String
    Dim event_date As String
    Select Case True
        Case racing_date_ Like "*札幌*"
            venue = "01"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "札幌")), "日目", "")
        Case racing_date_ Like "*函館*"
            venue = "02"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "函館")), "日目", "")
        Case racing_date_ Like "*福島*"
            venue = "03"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "福島")), "日目", "")
        Case racing_date_ Like "*新潟*"
            venue = "04"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "新潟")), "日目", "")
        Case racing_date_ Like "*東京*"
            venue = "05"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "東京")), "日目", "")
        Case racing_date_ Like "*中山*"
            venue = "06"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "中山")), "日目", "")
        Case racing_date_ Like "*中京*"
            venue = "07"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "中京")), "日目", "")
        Case racing_date_ Like "*京都*"
            venue = "08"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "京都")), "日目", "")
        Case racing_date_ Like "*阪神*"
            venue = "09"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "阪神")), "日目", "")
        Case racing_date_ Like "*小倉*"
            venue = "10"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "小倉")), "日目", "")
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
    Dim base_url_parameters As Object
    Set base_url_parameters = CreateObject("System.Collections.ArrayList")
    
    base_url_parameters.Add (convertRacingDate(racing_date.Value))
    
    Do
        Set next_racing_date = Cells.FindNext(After:=next_racing_date)
        
        If next_racing_date Is Nothing Then
            Exit Do
        End If
 
        If next_racing_date.Address = racing_date.Address Then
            Exit Do
        End If
        
        base_url_parameters.Add (convertRacingDate(next_racing_date.Value))
    Loop

    Set makeBaseURLParameters = base_url_parameters

End Function

Private Function makeURLParameters(base_url_parameters_ As Object) As Integer

    Set url_parameters = CreateObject("System.Collections.ArrayList")
    
    Dim i As Integer: i = 0
    Dim j As Integer: j = 1
    For i = 0 To base_url_parameters_.Count - 1
        For j = 1 To LAST_RACE_NUMBER
            url_parameters.Add (Year(Now) & base_url_parameters_(i) & Format(j, "00"))
        Next
    Next
    
    makeURLParameters = url_parameters.Count

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
    
    Set race_dates = CreateObject("System.Collections.ArrayList")
    Dim nextRow As Integer: nextRow = 1
    Do
        If InStr(Sheets(1).Cells(title_race_date.Cells.Row + nextRow, 1).Value, "/") = 0 Then
            Exit Do
        Else
            race_dates.Add (Sheets(1).Cells(title_race_date.Cells.Row + nextRow, 1).Value)
            nextRow = nextRow + 1
        End If
    Loop

    Set defineRaceDates = race_dates
    Exit Function
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function

Public Function fetchRacePrediction(url_netkeiba_prediction_parameter As String)

    With ThisWorkbook.Sheets(2).QueryTables.Add(Connection:= _
        "URL;https://race.netkeiba.com/?pid=yoso&id=c" & url_netkeiba_prediction_parameter, Destination:=ThisWorkbook.Sheets(2).Range("$A$1"))
        .Name = "url_netkeiba_prediction_parameter"
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

