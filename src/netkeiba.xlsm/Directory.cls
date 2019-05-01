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

Private Sub Class_Initialize()
    
End Sub

Private Function createFolder()

    If Dir(output_path_ & "\" & folder_name_, vbDirectory) = "" Then
        MkDir output_path_ & "\" & folder_name_
        createFolder = 0
        Exit Function
    End If
    
    createFolder = 1

End Function

Private Function createFile(file_name As String) As Long
    
    Dim file_system_object As Object
    Set file_system_object = CreateObject("Scripting.FileSystemObject")
    
    Dim file_path As String: file_path = output_path_ & "\" & folder_name_ & "\" & Replace(file_name, "/", "") & ".xlsx"

    If Not file_system_object.FileExists(file_path) Then
        Dim file As Object
        Set file = Workbooks.Add
        file.SaveAs file_path
        file.Close
        Set file = Nothing
        
        
        
        
        createFile = 0
        Exit Function
    End If
    
    createFile = 1
    
End Function

Public Function cofigureOutputPath(ByVal output_path As String)
On Error GoTo ErrorHandler

    output_path_ = output_path
    
Exit Function
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function

Public Function createFiles(ByVal folder_name As String, ByRef file_names() As String) As Boolean
On Error GoTo ErrorHandler

    folder_name_ = folder_name
    file_names_() = file_names()
    
    If createFolder = 0 Then
        Debug.Print "フォルダを作成しました"
    End If
    
    Dim i As Integer
    For i = 1 To UBound(file_names_())
        If createFile(file_names_(i)) = 0 Then
            Debug.Print ("開催日:" & file_names(i) & "のファイルを作成しました")
        End If
    Next
    
createFiles = True
Exit Function
ErrorHandler:
    MsgBox Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function
