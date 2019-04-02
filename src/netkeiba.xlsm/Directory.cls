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

Private output_path As String
Private folder_name As String
Private file_names As Object

Private Sub Class_Initialize()
    
End Sub

Private Function createFolder()

    If Dir(output_path & "\" & folder_name, vbDirectory) = "" Then
        MkDir output_path & "\" & folder_name
    End If

End Function

Private Function createFile(folder_name, file_ame_)
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim path As String
    path = output_path & folderPath_ & "\" & fileName_ & ".xlsx"

    If Not FSO.FileExists(path) Then
        Dim file As Workfile
        Set file = Workfiles.Add
        file.SaveAs basePath & "\" & folderPath_ & "\" & fileName_
        file.Close
        Set file = Nothing
    End If
    
End Function

Public Function cofigureOutputPath(output_path_ As String)
    output_path = output_path_
End Function

Public Function createFiles(folder_name_ As String, file_names_ As Object) As Boolean
On Error GoTo ErrorHandler

    folder_name = folder_name_
    Set file_names = file_names_
    createFolder
    
    Dim i As Integer
    For i = 0 To file_names.Count - 1
        Debug.Print file_names(i)
'       createFile (file_names(i))
    Next
createFiles = True
Exit Function
: ErrorHandler
End Function
