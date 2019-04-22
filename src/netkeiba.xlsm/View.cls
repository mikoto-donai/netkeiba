VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Sub Class_Initialize()
    ThisWorkbook.Sheets(1).Cells.Clear
End Sub
Public Function fetchItems(url As String) As String()

    With ThisWorkbook.Sheets(1).QueryTables.Add( _
         Connection:="URL;" & url, _
         destination:=Range("$A$1"))
        .Name = url
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
    
    Dim fetched_items_size As Long
    fetched_items_size = ThisWorkbook.Sheets(1).UsedRange.Rows.Count
    
    Dim fetched_items() As String
    ReDim fetched_items(fetched_items_size)
    
    Dim i As Long
    For i = 1 To fetched_items_size
        fetched_items(i) = ThisWorkbook.Sheets(1).Cells(i, 1).Value
    Next

    fetchItems = fetched_items()
    
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
