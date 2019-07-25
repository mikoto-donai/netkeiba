Attribute VB_Name = "Logger"
Option Explicit

Private start_time_ As Single   '�}�N���J�n����
Private end_time_ As Single     '�}�N���I������
Private log_path_ As String     '���O�o�͐�̃t���p�X
Private log As Object  '���O�I�u�W�F�N�g

Public Static Function initialize()
    
    start_time_ = Timer
    
    Dim file_system As Object
    Set file_system = CreateObject("Scripting.FileSystemObject")
    
    log_path_ = ThisWorkbook.Path & "\..\log\" & Format(Date, "yyyymmdd")
    If file_system.FileExists(log_path_) = False Then
        file_system.CreateTextFile log_path_
    End If

    Set log = file_system.OpenTextFile(log_path_, 8)
    Logger.writeLog "�������J�n���܂���"
    
End Function

Public Static Function writeLog(ByVal log_content As String)
    Logger.WriteLine Now & vbTab & log_content
End Function

Public Static Function finalize()

    end_time_ = Timer
    Logger.writeLog "�������������܂��� - ��Ǝ���: " & end_time_ - start_time_ & "�b" & vbTab

End Function
