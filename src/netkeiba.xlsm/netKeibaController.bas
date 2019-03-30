Attribute VB_Name = "netKeibaController"
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
