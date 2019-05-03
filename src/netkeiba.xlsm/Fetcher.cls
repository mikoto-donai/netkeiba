VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private url_ As String
Private target_sheet_ As Object
Private items_() As String

Private Sub Class_Initialize()
    Application.DisplayAlerts = False
End Sub

Public Property Let url(ByVal url As String)
    url_ = url
End Property

Public Property Let targetSheet(ByRef target_sheet As Object)
    Set target_sheet_ = target_sheet
End Property

Public Property Get items() As String()
    items = items_()
End Property

Public Function fetchItems()

    target_sheet_.Cells.Clear
    
    With target_sheet_.QueryTables.Add( _
         Connection:="URL;" & url_, _
         destination:=Range("$A$1"))
        .Name = url_
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
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
    
    Dim items_size As Long
    items_size = target_sheet_.UsedRange.Rows.Count
    
    ReDim items_(items_size)
    
    Dim i As Long
    For i = 1 To items_size
        items_(i) = target_sheet_.Cells(i, 1).Value
    Next
    
End Function

Private Function fetchRacePrediction(ByRef destination_sheet As Object)

Workbooks("netkeiba.xlsm").Sheets("fetched_data").Cells.Clear
Workbooks("netkeiba.xlsm").Activate


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

Function exporReportAsHTML()

    With ActiveWorkbook.PublishObjects.Add(xlSourceSheet, ActiveWorkbook.path & "\results.html", "Sheet1", "", xlHtmlStatic, "image")
        .Publish (True)
        .AutoRepublish = False
    End With

End Function

Function exportReportAsPDF()

    With Sheets("Sheet1").PageSetup
       .Orientation = xlLandscape
       .TopMargin = 0
       .LeftMargin = 0
    End With
     
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=ActiveWorkbook.path & "\results.pdf"

End Function

Private Sub Class_Terminate()
        Application.DisplayAlerts = True
End Sub
