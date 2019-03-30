VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RaceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Initialize()
    Sheets("Sheet1").Cells.Clear
End Sub

Public Function importRaceList()

    With ThisWorkbook.Sheets(1).QueryTables.Add(Connection:= _
        "URL;https://race.netkeiba.com/?pid=race_list", Destination:=Range( _
        "$A$1"))
        .Name = "?pid=race_list&id=c0331"
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
    
'    If ActiveSheet.Cells(1, 1).Value = "" Then
'       Debug.Print ("test")
'    End If


End Function

Private Function defineRaceDates()

    Dim title_race_date As Range
    Set title_race_date = Cells.Find(What:="開催日程")

    If title_race_date Is Nothing Then
        Exit Function
    Else
        Dim LAST_RACE_DATE As Integer: LAST_RACE_DATE = 5
        For i = 1 To LAST_RACE_DATE
            Debug.Print Sheets(1).Cells(title_race_date.Cells.Row + i, 1).Value
        Next
    End If

End Function

Private Function defineRaceInformation()

    Dim race_information As Range
    Set race_information = Cells.Find(What:="*回*日目")
   
    If race_information Is Nothing Then
        Exit Function
    End If
    
    Dim next_race_information As Range: Set next_race_information = race_information
    createURLParameter (race_information.Value)
    
    Do
        Set next_race_information = Cells.FindNext(After:=next_race_information)
        
        If next_race_information Is Nothing Then
            Exit Do
        End If
 
        If next_race_information.Address = race_information.Address Then
            Exit Do
        End If
        
        createURLParameter (next_race_information.Value)
    Loop

End Function
Private Function createURLParameter(race_information_ As String)
    
    Dim times As String
    times = Left(race_information_, InStr(race_information_, "回") - 1)
    
    Dim formatted_times As String
    formatted_times = Format(times, "00")

    Dim venue As String
    Dim event_date As String

    Select Case True
        Case race_information_ Like "*札幌*"
            venue = "01"
            event_date = Replace(Right(race_information_, InStr(race_information_, "札幌")), "日目", "")
        Case race_information_ Like "*函館*"
            venue = "02"
            event_date = Replace(Right(race_information_, InStr(race_information_, "函館")), "日目", "")
        Case race_information_ Like "*福島*"
            venue = "03"
            event_date = Replace(Right(race_information_, InStr(race_information_, "福島")), "日目", "")
        Case race_information_ Like "*新潟*"
            venue = "04"
            event_date = Replace(Right(race_information_, InStr(race_information_, "新潟")), "日目", "")
        Case race_information_ Like "*東京*"
            venue = "05"
            event_date = Replace(Right(race_information_, InStr(race_information_, "東京")), "日目", "")
        Case race_information_ Like "*中山*"
            venue = "06"
            event_date = Replace(Right(race_information_, InStr(race_information_, "中山")), "日目", "")
        Case race_information_ Like "*中京*"
            venue = "07"
            event_date = Replace(Right(race_information_, InStr(race_information_, "中京")), "日目", "")
        Case race_information_ Like "*京都*"
            venue = "08"
            event_date = Replace(Right(race_information_, InStr(race_information_, "京都")), "日目", "")
        Case race_information_ Like "*阪神*"
            venue = "09"
            event_date = Replace(Right(race_information_, InStr(race_information_, "阪神")), "日目", "")
        Case race_information_ Like "*小倉*"
            venue = "10"
            event_date = Replace(Right(race_information_, InStr(race_information_, "小倉")), "日目", "")
        Case Else
            Err.Raise number:=513, Description:="開催場所の変換でエラーが発生しました"
    End Select
    
    Dim formatted_event_date As String
    formatted_event_date = Format(event_date, "00")

    createURLParameter = venue & formatted_times & formatted_event_date
    Debug.Print createURLParameter
    
End Function

Public Function defineURLs()
    On Error GoTo ErrorHandler

    defineRaceDates
    defineRaceInformation
    
    Exit Function
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function

