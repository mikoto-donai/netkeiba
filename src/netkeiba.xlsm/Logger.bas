Attribute VB_Name = "Logger"
Option Explicit

Private start_time_ As Single   'マクロ開始時間
Private end_time_ As Single     'マクロ終了時間
Private log_path_ As String     'ログ出力先のフルパス
Private log As Object  'ログオブジェクト

Public Static Function initialize()
    
    start_time_ = Timer
    
    Dim file_system As Object
    Set file_system = CreateObject("Scripting.FileSystemObject")
    
    log_path_ = ThisWorkbook.Path & "\..\log\" & Format(Date, "yyyymmdd")
    If file_system.FileExists(log_path_) = False Then
        file_system.CreateTextFile log_path_
    End If

    Set log = file_system.OpenTextFile(log_path_, 8)
    Logger.writeLog "処理を開始しました"
    
End Function

Public Static Function writeLog(ByVal log_content As String)
    Logger.WriteLine Now & vbTab & log_content
End Function

Public Static Function finalize()

    end_time_ = Timer
    Logger.writeLog "処理を完了しました - 作業時間: " & end_time_ - start_time_ & "秒" & vbTab

End Function
