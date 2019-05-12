VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RaceEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private race_events_() As String     '開催日　　ex. 04/13(土), 04/14(日), 04/20(土), 04/21(日),
Private topic_race_event_ As String     '注目レース開催日　　ex. 04/13(土)
Private race_events_parameters_() As String       '開催日のURLパラメータ　　ex. p0420, c0421, n0422,
Private current_race_events_()  As String   '注目レース開催日を起点とした直近の開催日     ex. 04/20(土), 04/21(日),
Private current_race_event_parameters_() As String       '注目レース開催日を起点とした直近の開催日パラメータ    ex. c0420, c0421,

Public Property Get event_dates() As String()
    event_dates = race_events_()
End Property
Public Property Get currentRaceEvents() As String()
    currentRaceEvents = current_race_events_()
End Property

Public Property Get current_race_event_parameters() As String()
    current_race_event_parameters = current_race_event_parameters_()
End Property

Public Function analyzeItems(ByRef items As Variant)
    
    If createRaceEvents(items) = -1 Then
        Err.Raise Number:=513, Description:="開催日の作成に失敗しました"
        Exit Function
    End If
    
    If createTopicRaceEvent(items) = -1 Then
        Err.Raise Number:=513, Description:="注目レース開催日の作成に失敗しました"
        Exit Function
    End If

    If createRaceEventParameters() = -1 Then
        Err.Raise Number:=513, Description:="開催日のURLパラメータの作成に失敗しました"
        Exit Function
    End If

    If createCurrentRaceEventParameters() = -1 Then
        Err.Raise Number:=513, Description:="注目レース開催日を起点とした直近の開催日パラメータ作成に失敗しました"
        Exit Function
    End If
    
    analyzeItems = 0

End Function

Private Function createTopicRaceEvent(ByRef items As Variant) As Long

    Dim i As Long
    For i = LBound(items, 1) To UBound(items, 1)
        If items(i, 1) = "今週の注目レース" And InStr(items(i + 2, 1), "/") > 0 Then
            topic_race_event_ = Left(items(i + 2, 1), 8)
            createTopicRaceEvent = 0
            Exit Function
        End If
    Next
    
    createTopicRaceEvent = 0

End Function

Private Function createRaceEvents(ByRef items As Variant) As Long
   
    ReDim race_events_(0)
    
    Dim i As Long
    For i = LBound(items, 1) To UBound(items, 1)
        If items(i, 1) = "開催日程" Then
            Do
                If InStr(items(i + UBound(race_events_()) + 1, 1), "/") = 0 Then
                    ReDim Preserve race_events_(UBound(race_events_()) - 1)
                    Exit Do
                End If
                race_events_(UBound(race_events_())) = items(i + UBound(race_events_()) + 1, 1)
                ReDim Preserve race_events_(UBound(race_events_()) + 1)
            Loop
            Exit For
        End If
    Next
    
    createRaceEvents = 0
    
End Function

Private Function createRaceEventParameters() As Long
    
    ReDim race_events_parameters_(UBound(race_events_()))
    ReDim current_race_events_(0)
    
    Dim i As Long
    For i = 0 To UBound(race_events_())
        race_events_parameters_(i) = convertRaceEventParameter(race_events_(i))
    Next
       
    If UBound(current_race_events_()) > 0 Then
        ReDim Preserve current_race_events_(UBound(current_race_events_()) - 1)
    End If
    
    createRaceEventParameters = 0

End Function

Private Function convertRaceEventParameter(ByVal race_event As String) As String

     Dim converted_race_event As String
     Dim converted_topic_race_event As String
     
     converted_race_event = Left(race_event, 5)
     converted_topic_race_event = Left(topic_race_event_, 5)
     
     If DateDiff("d", CDate(converted_topic_race_event), CDate(converted_race_event)) >= 5 Then
        converted_race_event = "n" & converted_race_event
        
     ElseIf DateDiff("d", CDate(converted_topic_race_event), CDate(converted_race_event)) < 5 _
     And DateDiff("d", CDate(converted_topic_race_event), CDate(converted_race_event)) >= -2 Then
     
        current_race_events_(UBound(current_race_events_())) = race_event
        ReDim Preserve current_race_events_(UBound(current_race_events_()) + 1)
        converted_race_event = "c" & converted_race_event
 
     ElseIf DateDiff("d", CDate(converted_topic_race_event), CDate(converted_race_event)) < -2 Then
        converted_race_event = "p" & converted_race_event
        
     Else
        converted_race_event = "n" & converted_race_event
     End If
     
     convertRaceEventParameter = Left(Replace(converted_race_event, "/", ""), 5)
     
End Function

Private Function createCurrentRaceEventParameters() As Long

    ReDim current_race_event_parameters_(0)
    Dim i As Long
    For i = 0 To UBound(race_events_())
        If Left(race_events_parameters_(i), 1) = "c" Then
            current_race_event_parameters_(UBound(current_race_event_parameters_())) = race_events_parameters_(i)
            ReDim Preserve current_race_event_parameters_(UBound(current_race_event_parameters_()) + 1)
        End If
    Next
    
    If UBound(current_race_event_parameters_()) > 0 Then
        ReDim Preserve current_race_event_parameters_(UBound(current_race_event_parameters_()) - 1)
    End If
    
    createCurrentRaceEventParameters = 0
    
End Function


