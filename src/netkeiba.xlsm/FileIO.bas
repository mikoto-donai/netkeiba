Attribute VB_Name = "FileIO"
Option Explicit

Private start_time_ As Single   'マクロ開始時間
Private end_time_ As Single     'マクロ終了時間
Private log_path_ As String     'ログ出力先のフルパス
Private log As Object  'ログオブジェクト

Public Static Function startLog()
    
    start_time_ = Timer
    
    Dim file_system As Object
    Set file_system = CreateObject("Scripting.FileSystemObject")
    
    log_path_ = ThisWorkbook.Path & "\..\log\" & Format(Date, "yyyymmdd")
    If file_system.FileExists(log_path_) = False Then
        file_system.CreateTextFile log_path_
    End If

    Set log = file_system.OpenTextFile(log_path_, 8)
    FileIO.writeLog "処理を開始しました"
    
End Function

Public Static Function writeLog(ByVal log_content As String)
    log.writeLine Now & vbTab & log_content
End Function

Public Static Function stopLog()

    end_time_ = Timer
    FileIO.writeLog "処理を完了しました - 作業時間: " & end_time_ - start_time_ & "秒" & vbTab

End Function

Public Static Function readFile(ByVal LINE_NUMBER As Long) As String

    Dim file_system_object As Object
    Set file_system_object = CreateObject("Scripting.FileSystemObject")
    
    Dim file_path As String: file_path = ThisWorkbook.Path & "\..\user\user"
    Dim text_stream As Object
    Const READING = 1
    Set text_stream = file_system_object.OpenTextFile(file_path, READING)
    
    Dim file_contents() As String: ReDim file_contents(0)
    Do Until text_stream.AtEndOfStream = True
        file_contents(UBound(file_contents())) = text_stream.ReadLine
        ReDim Preserve file_contents(UBound(file_contents()) + 1)
    Loop
    
    ReDim Preserve file_contents(UBound(file_contents()) - 1)
    text_stream.Close
    readFile = file_contents(LINE_NUMBER - 1)
    
End Function

