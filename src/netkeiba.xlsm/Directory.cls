VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private output_path_ As String
Private folder_name_ As String
Private file_names_() As String
Private contents_() As String

Private Sub Class_Initialize()
    
End Sub

Public Property Let outputPath(ByVal output_path As String)
    output_path_ = output_path
End Property

Public Property Let folderName(ByVal folder_name As String)
    folder_name_ = folder_name
End Property

Public Property Let fileNames(ByRef file_names() As String)
    file_names_() = file_names()
End Property

Public Function createFiles()

    If createFolder() <> 0 Then
    End If
    
    If createExcelFiles() <> 0 Then
    End If

End Function

Private Function createFolder() As Long

    If Dir(output_path_ & "\" & folder_name_, vbDirectory) = "" Then
        MkDir output_path_ & "\" & folder_name_
        createFolder = 0
        Exit Function
    End If
    
    createFolder = 0

End Function

Private Function createExcelFiles() As Long
    
    Dim file_system_object As Object
    Set file_system_object = CreateObject("Scripting.FileSystemObject")

    Dim file_path As String
    Dim b_book As New Book
    Dim i As Long
    For i = LBound(file_names_()) To UBound(file_names_())
        file_path = output_path_ & "\" & folder_name_ & "\" & Replace(file_names_(i), "/", "") & ".xlsx"
        
        If Not file_system_object.FileExists(file_path) Then
            Dim file As Object
            Set file = Workbooks.Add
            b_book.targetSheet = file.Worksheets(1)
            b_book.url = "https://race.netkeiba.com/?pid=top"
            b_book.fetchItems
            file.SaveAs file_path
            file.Close
            Set file = Nothing
        End If
    Next

    createExcelFiles = 0
    
End Function

