VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PastRace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private race_dates_() As String  '���[�X��  ex. race_date(0) / 2�񓌋�7����
Private race_parameters_() As String   '�e���[�X��URL�p�����[�^  ex.  race_parameters_(0) / 2019050207
Private file_names_() As String  '�Z���l�i�[��̃t�@�C���� ex. file_names(0,1) / 2�񓌋�7����:2R
Private race_results_() As Variant    '�Z���l  ex.  race_results_(0, 1) / 2�񓌋�7����:2R �̃Z���l

Public Property Get fileNames() As String()
    fileNames = file_names_()
End Property

Public Property Get raceResults() As Variant()
    raceResults = race_results_()
End Property

Public Function analyzeItems(ByVal race_year As Long, ByRef items As Variant)
    
    If createRaceDates(items) = -1 Then
        Err.Raise Number:=513, Description:="���[�X���̍쐬�Ɏ��s���܂���"
        Exit Function
    End If
    
    If createRaceParameters(race_year) = -1 Then
        Err.Raise Number:=513, Description:="���[�X�p�����[�^�̍쐬�Ɏ��s���܂���"
        Exit Function
    End If
    
    If fetchRaceResults() = -1 Then
        Err.Raise Number:=513, Description:="�e���[�X�̎擾�Ɏ��s�f���܂���"
        Exit Function
    End If

End Function

Private Function createRaceDates(ByRef items As Variant) As Long

    Dim i As Long
    ReDim race_dates_(0)
    
    For i = LBound(items, 1) To UBound(items, 1)
        If InStr(items(i, 1), "��") > 0 Then
            race_dates_(UBound(race_dates_)) = Mid(items(i, 1), InStr(items(i, 1), " ") + 1) & "��"
            ReDim Preserve race_dates_(UBound(race_dates_) + 1)
        End If
    Next
    
    ReDim Preserve race_dates_(UBound(race_dates_) - 1)
    
    createRaceDates = 0

End Function

Private Function createRaceParameters(ByVal race_year As Long) As Long

    Dim i As Long
    ReDim race_parameters_(0)
    
    For i = LBound(race_dates_()) To UBound(race_dates_())
        race_parameters_(UBound(race_parameters_)) = CStr(race_year) & convertRaceDateIntoParameter(race_dates_(i))
        ReDim Preserve race_parameters_(UBound(race_parameters_) + 1)
    Next
    
    ReDim Preserve race_parameters_(UBound(race_parameters_) - 1)
    
    createRaceParameters = 0
    
End Function

Private Function convertRaceDateIntoParameter(ByVal race_date As String) As String
    
    Dim event_time As String
    event_time = Format(Left(race_date, 1), "00")

    Dim event_venue As String
    Dim event_date As String
    
    Select Case True
        Case race_date Like "*�D�y*"
            event_venue = "01"
            event_date = Mid(race_date, InStr(race_date, "�D�y") + 2)
        Case race_date Like "*����*"
            event_venue = "02"
            event_date = Mid(race_date, InStr(race_date, "����") + 2)
        Case race_date Like "*����*"
            event_venue = "03"
            event_date = Mid(race_date, InStr(race_date, "����") + 2)
        Case race_date Like "*�V��*"
            event_venue = "04"
            event_date = Mid(race_date, InStr(race_date, "�V��") + 2)
        Case race_date Like "*����*"
            event_venue = "05"
            event_date = Mid(race_date, InStr(race_date, "����") + 2)
        Case race_date Like "*���R*"
            event_venue = "06"
            event_date = Mid(race_date, InStr(race_date, "���R") + 2)
        Case race_date Like "*����*"
            event_venue = "07"
            event_date = Mid(race_date, InStr(race_date, "����") + 2)
        Case race_date Like "*���s*"
            event_venue = "08"
            event_date = Mid(race_date, InStr(race_date, "���s") + 2)
        Case race_date Like "*��_*"
            event_venue = "09"
            event_date = Mid(race_date, InStr(race_date, "��_") + 2)
        Case race_date Like "*���q*"
            event_venue = "10"
            event_date = Mid(race_date, InStr(race_date, "���q") + 2)
        Case Else
            Err.Raise Number:=513, Description:="���[�X���̕ϊ��ŃG���[���������܂���"
    End Select
 
    event_date = Format(Replace(event_date, "����", ""), "00")

    convertRaceDateIntoParameter = event_venue & event_time & event_date
    
End Function

Private Function fetchRaceResults() As Long

    ReDim race_results_(UBound(race_parameters_()), RACE_NUMBER.LAST - 1)
    ReDim file_names_(UBound(race_parameters_()), RACE_NUMBER.LAST - 1)
    
    Dim i As Integer
    Dim j As Integer
    
    Dim o_fetcher As New Fetcher
    
    For i = LBound(race_parameters_()) To UBound(race_parameters_())
        For j = RACE_NUMBER.FIRST - 1 To RACE_NUMBER.LAST - 1
            
            o_fetcher.url = "https://db.netkeiba.com/race/" & race_parameters_(i) & Format(j + 1, "00")
            o_fetcher.targetSheet = ThisWorkbook.Sheets("fetched_data")
            o_fetcher.fetchItems 1
            
            If Not IsEmpty(o_fetcher.items) Then
                file_names_(i, j) = race_dates_(i) & ":" & j + 1 & "R"
                race_results_(i, j) = o_fetcher.items
                StaticModule.logContent file_names_(i, j) & "���擾��"
            End If
            
        Next
            If i = 1 Then
                Exit For
            End If
    Next
    
    fetchRaceResults = 0
    
End Function