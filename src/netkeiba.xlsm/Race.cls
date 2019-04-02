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
    times = Left(racing_date_, InStr(racing_date_, "��") - 1)
    
    Dim formatted_times As String
    formatted_times = Format(times, "00")

    Dim venue As String
    Dim event_date As String
    Select Case True
        Case racing_date_ Like "*�D�y*"
            venue = "01"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "�D�y")), "����", "")
        Case racing_date_ Like "*����*"
            venue = "02"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "����")), "����", "")
        Case racing_date_ Like "*����*"
            venue = "03"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "����")), "����", "")
        Case racing_date_ Like "*�V��*"
            venue = "04"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "�V��")), "����", "")
        Case racing_date_ Like "*����*"
            venue = "05"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "����")), "����", "")
        Case racing_date_ Like "*���R*"
            venue = "06"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "���R")), "����", "")
        Case racing_date_ Like "*����*"
            venue = "07"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "����")), "����", "")
        Case racing_date_ Like "*���s*"
            venue = "08"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "���s")), "����", "")
        Case racing_date_ Like "*��_*"
            venue = "09"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "��_")), "����", "")
        Case racing_date_ Like "*���q*"
            venue = "10"
            event_date = Replace(Right(racing_date_, InStr(racing_date_, "���q")), "����", "")
        Case Else
            Err.Raise number:=513, Description:="���[�X���̕ϊ��ŃG���[���������܂���"
    End Select
    
    Dim formatted_event_date As String
    formatted_event_date = Format(event_date, "00")

    convertRacingDate = venue & formatted_times & formatted_event_date

End Function

Private Function makeBaseURLParameters() As Object

    Dim racing_date As Range
    Set racing_date = Cells.Find(What:="*��*����")
   
    If racing_date Is Nothing Then
         Err.Raise number:=513, Description:="���[�X�����擾�ł��܂���ł���"
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
        Err.Raise number:=513, Description:="URL���쐬�ł��܂���ł���"
    End If

    Exit Function
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "�G���["
End Function

Public Function defineRaceDates() As Object
    On Error GoTo ErrorHandler

    Dim title_race_date As Range
    Set title_race_date = Cells.Find(What:="�J�Ó���")

    If title_race_date Is Nothing Then
        Err.Raise number:=513, Description:="�J�Ó���������܂���"
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
    Debug.Print race_dates(3)
    Set defineRaceDates = race_dates
    
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "�G���["
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
