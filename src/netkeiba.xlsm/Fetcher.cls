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
Private items_ As Variant    'リモートから取得してセルに展開した内容
Private Sub Class_Initialize()
    Me.targetSheet = ThisWorkbook.Sheets("fetched_data")
    Me.url = "https://race.netkeiba.com"
End Sub

Public Property Let url(ByVal url As String)
    url_ = url
End Property

Public Property Let targetSheet(ByRef target_sheet As Object)
    Set target_sheet_ = target_sheet
End Property

Public Property Get items() As Variant
    items = items_
End Property

Public Function fetchItems(Optional ByVal web_selection_type As String = xlEntirePage)
    
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
        .WebSelectionType = web_selection_type
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
    items_ = target_sheet_.UsedRange

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
