Attribute VB_Name = "netKeibaController"
Function initializeBook()

   Sheets("Sheet1").Cells.Clear

   
End Function

Function importData()

    With ActiveSheet.QueryTables.Add(Connection:="URL;https://race.netkeiba.com/?pid=yoso&id=p201906030101", Destination:=Range("$A$1"))
        .name = "?kd=1&tm=d&vl=a&mk=1&p=1"
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
        .WebSelectionType = xlSpecifiedTables
        .WebTables = "1"
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
    If ActiveSheet.Cells(1, 1).Value = "" Then
       Debug.Print ("test")
    End If

End Function
 
Function convertData()

    

End Function

Function exporDataAsHTML()

    With ActiveWorkbook.PublishObjects.Add(xlSourceSheet, ActiveWorkbook.path & "\results.html", "Sheet1", "", xlHtmlStatic, "image")
        .Publish (True)
        .AutoRepublish = False
    End With

End Function

Function exportDataAsPDF()

    With Sheets("Sheet1").PageSetup
       .Orientation = xlLandscape
       .TopMargin = 0
       .LeftMargin = 0
    End With
     
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=ActiveWorkbook.path & "\results.pdf"

End Function
