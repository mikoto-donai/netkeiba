Attribute VB_Name = "StaticModule"
Option Explicit

Private start_time_ As Single   'マクロ開始時間
Private end_time_ As Single     'マクロ終了時間
Private log_path_ As String     'ログ出力先のフルパス
Private log_contents_() As String   '出力するログの中身

Public Static Function initialize()

    Application.DisplayAlerts = False
    start_time_ = Timer
    log_path_ = ThisWorkbook.Path & "\..\log\" & Format(Date, "yyyymmdd")
    ReDim log_contents_(0)
    log_contents_(0) = Now & vbTab & "処理を開始しました"
    
End Function

Public Static Function logContent(ByVal log_content As String)
    
    ReDim Preserve log_contents_(UBound(log_contents_) + 1)
    log_contents_(UBound(log_contents_)) = log_content
    
End Function

Public Static Function finalize()

    end_time_ = Timer
    
    Dim file_system As Object
    Set file_system = CreateObject("Scripting.FileSystemObject")

    If file_system.FileExists(log_path_) = False Then
        file_system.CreateTextFile log_path_
    End If
    
    Dim log As Object
    Set log = file_system.OpenTextFile(log_path_, 8)

    Dim i As Long
    For i = LBound(log_contents_()) To UBound(log_contents_())
        log.WriteLine log_contents_(i) & vbTab
    Next
    
    log.WriteLine Now & vbTab & "処理を完了しました - 作業時間: " & end_time_ - start_time_ & "秒" & vbTab
    Application.DisplayAlerts = True

End Function
