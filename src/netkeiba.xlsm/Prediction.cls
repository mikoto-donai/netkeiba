VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Prediction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum RACE_NUMBER
    FIRST = 1
    LAST = 12
End Enum

Private prediction_parameters_() As Variant   '�e���[�X�̋��n�\�zURL�p�����[�^: prediction_parameters_(race_date, race_number)  / ex.  prediction_parameters_(0, 10) : 201905020701
Private prediction_race_dates_() As String    'pid=yoso&id={prediction_parameter} �œ����鋣�n�\�z�f�[�^: predictions_targets(race_date, race_number)  /  ex.  predictions_targets(0, 10) : 2�񓌋�7����:11R
Private predictions_() As Variant    'pid=yoso&id={prediction_parameter} �œ����鋣�n�\�z�f�[�^: predictions_(race_date, race_number)  /  ex.  predictions_(0, 10) : 2�񓌋�7����:11R �̗\�z�f�[�^Range�I�u�W�F�N�g

Public Property Get predictionRaceDates() As String()
    predictionRaceDates = prediction_race_dates_()
End Property

Public Property Get predictions() As Variant()
    predictions = predictions_()
End Property

Public Function analyzeItems(ByRef current_race_dates() As String)
    
    If createPredictionParameters(current_race_dates()) <> 0 Then
    End If
    
    If fetchPredictions(current_race_dates()) <> 0 Then
    End If

End Function

Private Function createPredictionParameters(ByRef current_race_dates() As String) As Long

    ReDim prediction_parameters_(UBound(current_race_dates()), RACE_NUMBER.LAST - 1)
    
    Dim i As Long
    Dim j As Long
    For i = LBound(current_race_dates()) To UBound(current_race_dates())
        For j = 0 To RACE_NUMBER.LAST - 1
            prediction_parameters_(i, j) = (Year(Now) & convertRaceDateIntoParameters(current_race_dates(i)) & Format(j + 1, "00"))
        Next
    Next
   
    createPredictionParameters = 0

End Function

Private Function convertRaceDateIntoParameters(ByVal current_race_date As String) As String
    
    Dim times As String
    times = Left(current_race_date, InStr(current_race_date, "��") - 1)
    
    Dim formatted_times As String
    formatted_times = Format(times, "00")

    Dim venue As String
    Dim event_date As String
    Select Case True
        Case current_race_date Like "*�D�y*"
            venue = "01"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "�D�y")), "����", "")
        Case current_race_date Like "*����*"
            venue = "02"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "����")), "����", "")
        Case current_race_date Like "*����*"
            venue = "03"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "����")), "����", "")
        Case current_race_date Like "*�V��*"
            venue = "04"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "�V��")), "����", "")
        Case current_race_date Like "*����*"
            venue = "05"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "����")), "����", "")
        Case current_race_date Like "*���R*"
            venue = "06"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "���R")), "����", "")
        Case current_race_date Like "*����*"
            venue = "07"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "����")), "����", "")
        Case current_race_date Like "*���s*"
            venue = "08"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "���s")), "����", "")
        Case current_race_date Like "*��_*"
            venue = "09"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "��_")), "����", "")
        Case current_race_date Like "*���q*"
            venue = "10"
            event_date = Replace(Right(current_race_date, InStr(current_race_date, "���q")), "����", "")
        Case Else
            Err.Raise number:=513, Description:="���[�X���̕ϊ��ŃG���[���������܂���"
    End Select
    
    Dim formatted_event_date As String
    formatted_event_date = Format(event_date, "00")

    convertRaceDateIntoParameters = venue & formatted_times & formatted_event_date

End Function

Private Function fetchPredictions(ByRef current_race_dates() As String) As Long

    ReDim prediction_race_dates_(UBound(prediction_parameters_, 1), UBound(prediction_parameters_, 2))
    ReDim predictions_(UBound(prediction_parameters_, 1), UBound(prediction_parameters_, 2))

    Dim i As Integer
    Dim j As Integer
    Dim o_fetcher As New fetcher
    Const web_selection_type As String = xlAllTables
    
    For i = LBound(prediction_parameters_, 1) To UBound(prediction_parameters_, 1)
'    If i = 1 Then
'        Exit For
'    End If
        For j = LBound(prediction_parameters_, 2) To UBound(prediction_parameters_, 2)
            
            o_fetcher.url = "https://race.netkeiba.com/?pid=yoso&id=c" & prediction_parameters_(i, j)
            o_fetcher.targetSheet = ThisWorkbook.Sheets("fetched_data")
            o_fetcher.fetchItems web_selection_type
            
            If Not IsEmpty(o_fetcher.items) Then
                predictions_(i, j) = o_fetcher.items
            End If
            
            prediction_race_dates_(i, j) = current_race_dates(i) & ":" & j + 1 & "R"
            
        Next
    Next
    
    fetchPredictions = 0
    
End Function


