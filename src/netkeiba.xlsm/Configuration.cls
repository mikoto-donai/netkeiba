VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private start_time_ As Single   '�}�N���J�n����
Private end_time_ As Single     '�}�N���I������
Private log_path_ As String     '���O�o�͐�̃t���p�X
Private log_contents_() As String   '�o�͂��郍�O�̒��g

Private Sub Class_Initialize()

    Application.DisplayAlerts = False
    start_time_ = Timer
    log_path_ = ThisWorkbook.Path & "\..\log\" & Format(Date, "yyyymmdd")
    ReDim log_contents_(0)
    log_contents_(0) = Now & vbTab & "�������J�n���܂���"
    
End Sub

Public Property Let logContet(ByVal log_content As String)
    
    ReDim Preserve log_contents_(UBound(log_contents_) + 1)
    log_contents_(UBound(log_contents_)) = log_content
    
End Property

Public Function finalize()

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
    
    log.WriteLine Now & vbTab & "�������������܂��� - ��Ǝ���: " & end_time_ - start_time_ & "�b" & vbTab
    Application.DisplayAlerts = True

End Function


