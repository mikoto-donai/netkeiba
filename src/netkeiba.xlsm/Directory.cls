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

Private output_path_ As String      '出力先
Private folder_name_ As String      '作成したファイルを格納するフォルダ名
Private file_names_() As String     'シート名_('ファイル名インデックス', 'シート名インデックス') 　　ex. file_names_(0,0) / ファイル名hoge:シート名fuga
Private contents_() As Variant    'Rangeオブジェクト_('ファイル名インデックス', 'シート名インデックス') 　　ex. contents_(0,0)  / ファイル0, シート0のRangeオブジェクト

Private Sub Class_Initialize()
    Me.outputPath = Environ("HOMEPATH") & "\" & "Desktop"
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

Public Property Let contents(ByRef contents() As Variant)
    contents_() = contents()
End Property

Public Function createFiles()

    If createFolder() = -1 Then
        Err.Raise Number:=513, Description:="フォルダの作成に失敗しました"
        Exit Function
    End If
    
    If createExcelFiles() = -1 Then
        Err.Raise Number:=513, Description:="ファイルの作成に失敗しました"
        Exit Function
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
    Application.SheetsInNewWorkbook = 1
    
    Dim file_path As String
    Dim file_name As String
    Dim i As Long
    Dim j As Long
    
    For i = LBound(file_names_(), 1) To UBound(file_names_(), 1)
        If file_names_(i, 0) = "" Then
            createExcelFiles = 1
            Exit Function
        End If
        
        file_name = Left(file_names_(i, 0), InStr(file_names_(i, 0), ":") - 1) & ".xlsx"
        file_path = output_path_ & "\" & folder_name_ & "\" & file_name
        
        If file_system_object.FileExists(file_path) Then
            file_system_object.DeleteFile file_path
        End If
        
        Dim file As Object
        Set file = Workbooks.Add
        
        For j = LBound(file_names_(), 2) To UBound(file_names_(), 2)
            If Not IsEmpty(contents_(i, j)) Then
                ActiveWorkbook.ActiveSheet.Cells(1, 1).Resize(UBound(contents_(i, j), 1), UBound(contents_(i, j), 2)).value = contents_(i, j)
                ActiveWorkbook.ActiveSheet.Name = Replace(file_names_(i, j), ":", "")
                ActiveWorkbook.Worksheets.Add after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
            End If
        Next
        
        ActiveWorkbook.Worksheets(1).Activate
        
        file.SaveAs file_path
        file.Close
        Set file = Nothing
        
    Next

    createExcelFiles = 0
    
End Function


