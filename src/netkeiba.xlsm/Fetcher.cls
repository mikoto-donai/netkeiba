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

Private url_ As String       '取得対象のURL
Private target_sheet_ As Object      '取得データの展開先
Private items_ As Variant    'リモートから取得して展開したセル値

Public Property Let url(ByVal url As String)
    url_ = url
End Property

Public Property Get items() As Variant
    items = items_
End Property

Public Function configureTargetSheet(ByVal targetSheetName As String)

    Dim isRenamed As Boolean: isRenamed = False
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.Name = targetSheetName Then
            isRenamed = True
            Exit For
        End If
    Next
    If Not isRenamed Then
        ThisWorkbook.Sheets(1).Name = targetSheetName
    End If
    
    Set target_sheet_ = ThisWorkbook.Sheets(targetSheetName)

End Function
 
Public Function fetchItems(Optional ByVal WEB_TABLE)
    
    target_sheet_.Activate
    target_sheet_.Cells.Clear
    
    Dim query_table As Object
    
    Set query_table = target_sheet_.QueryTables.Add( _
         Connection:="URL;" & url_, _
         Destination:=Range("$A$1"))
    
    With query_table
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
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
    End With
    
    With query_table
        If IsMissing(WEB_TABLE) Then
            .WebSelectionType = xlAllTables
        Else
            .WebSelectionType = xlSpecifiedTables
            .WebTables = WEB_TABLE
        End If
        .Refresh BackgroundQuery:=False
    End With
    
    If VarType(target_sheet_.UsedRange) = vbString Then
        Dim fakeArray(1, 1) As String
        fakeArray(1, 1) = target_sheet_.UsedRange
        items_ = fakeArray
        Exit Function
    End If
    
    items_ = target_sheet_.UsedRange

End Function

Private Sub Class_Terminate()
    target_sheet_.Cells.Clear
End Sub
