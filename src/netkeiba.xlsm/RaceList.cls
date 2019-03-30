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

    Dim titleOfRaceDate As Range
    Set titleOfRaceDate = Cells.Find(What:="開催日程")

    If titleOfRaceDate Is Nothing Then
        Exit Function
    Else
        Dim LAST_RACEDATE As Integer: LAST_RACEDATE = 5
        For i = 1 To LAST_RACEDATE
            Debug.Print Sheets(1).Cells(titleOfRaceDate.Cells.Row + i, 1).Value
        Next
    End If

End Function

Private Function defineRaceInformation()

    Dim raceInformation As Range
    Set raceInformation = Cells.Find(What:="*回*日目")
   
    If raceInformation Is Nothing Then
        Exit Function
    End If
    
    Dim nextRaceInformation As Range: Set nextRaceInformation = raceInformation
    
    Debug.Print createURLParameter(raceInformation.Value)
    
    Do
        Set nextRaceInformation = Cells.FindNext(After:=nextRaceInformation)
        
        If nextRaceInformation Is Nothing Then
            Exit Do
        End If
 
        If nextRaceInformation.Address = raceInformation.Address Then
            Exit Do
        End If
        
        Debug.Print createURLParameter(nextRaceInformation.Value)
    Loop

End Function
Private Function createURLParameter(raceInformation_ As String)

    Dim convertedRaceInformation As String
    
    Select Case True
        Case raceInformation_ Like "*札幌*"
            convertedRaceInformation = Replace(raceInformation_, "札幌", "01")
        Case raceInformation_ Like "*函館*"
            convertedRaceInformation = Replace(raceInformation_, "函館", "02")
        Case raceInformation_ Like "*福島*"
            convertedRaceInformation = Replace(raceInformation_, "福島", "03")
        Case raceInformation_ Like "*新潟*"
            convertedRaceInformation = Replace(raceInformation_, "新潟", "04")
        Case raceInformation_ Like "*東京*"
            convertedRaceInformation = Replace(raceInformation_, "東京", "05")
        Case raceInformation_ Like "*中山*"
            convertedRaceInformation = Replace(raceInformation_, "中山", "06")
        Case raceInformation_ Like "*中京*"
            convertedRaceInformation = Replace(raceInformation_, "中京", "07")
        Case raceInformation_ Like "*京都*"
            convertedRaceInformation = Replace(raceInformation_, "京都", "08")
        Case raceInformation_ Like "*阪神*"
            convertedRaceInformation = Replace(raceInformation_, "阪神", "09")
        Case raceInformation_ Like "*小倉*"
            convertedRaceInformation = Replace(raceInformation_, "小倉", "10")
        Case Else
            Err.Raise Number:=513, Description:="開催場所の変換でエラーが発生しました"
    End Select

    createURLParameter = convertedRaceInformation
    
End Function

Public Function defineURLs()
    On Error GoTo ErrorHandler

    defineRaceDates
    defineRaceInformation
    
    
    Exit Function
ErrorHandler:
    MsgBox Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function

