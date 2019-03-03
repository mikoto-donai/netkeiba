VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelBookHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private basePath As String
Private folderNames As Variant
Private bookNames As Variant

Private Sub Class_Initialize()
    basePath = ThisWorkbook.path
End Sub

Public Property Let setFolderNames(folderNames_ As Variant)
    folderNames = folderNames_
End Property

Public Property Let setBookNames(bookNames_ As Variant)
    bookNames = bookNames_
End Property

Public Function createFoldersAndBooks()
    
    Dim folderName As Variant
    Dim bookName As Variant
    
    For Each folderName In folderNames
        Call createFolder(folderName)
        
        For Each bookName In bookNames
            Call createBook(folderName, bookName)
        Next
    Next

End Function

Private Function createFolder(folderName_)
    
    Dim path As String
    path = basePath & "\" & folderName_

    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If

End Function

Private Function createBook(folderPath_, bookName_)
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim path As String
    path = basePath & folderPath_ & "\" & bookName_ & ".xlsx"

    If Not FSO.FileExists(path) Then
        Dim book As Workbook
        Set book = Workbooks.Add
        book.SaveAs basePath & "\" & folderPath_ & "\" & bookName_
        book.Close
        Set book = Nothing
    End If

End Function
