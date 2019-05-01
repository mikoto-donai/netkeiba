VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Race"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private race_lists_() As String     '04/20(“y), 04/21(“ú),
Private race_list_parameters_() As String       'p0420, c0421, n0422,
Private current_race_list_parameters_() As String       'c0420, c0421,
Private race_dates_() As String      '2‰ñ“Œ‹4“ú–Ú, 3‰ñ‹“s3“ú–Ú,
Private race_date_parameters_() As String      '201905020101, 201905020102,
Private current_race_names() As String    '0420(“y)2‰ñ“Œ‹4“ú–Ú, 0420(“y)3‰ñ‹“s3“ú–Ú, 0421(“ú)2‰ñ“Œ‹5“ú–Ú, 0421(“ú)3‰ñ‹“s6“ú–Ú
Private url_parameter_predictions_ As Object

Private Sub Class_Initialize()
    
End Sub

'Public Function extractEventInfo(ByRef fetched_items() As String) As String
'    On Error GoTo ErrorHandler
'
'    Dim i As Long
'    For i = 0 To UBound(fetched_items())
'        If fetched_items(i) = "ŠJÃî•ñ" Then
'            Dim event_info As String
'            event_info = Mid(fetched_items(i + 1), InStr(fetched_items(i + 1), "(") + 1, 5)
'
'            Dim formatted_event_info
'            formatted_event_info = Format(Trim(event_info), "mmdd")
'
'            Exit For
'        End If
'    Next
'
'    extractEventInfo = event_info
'
'    Exit Function
'ErrorHandler:
'    Debug.Print Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "ƒGƒ‰["
'End Function

Private Function createRaceLists(ByRef fetched_items() As String) As String()

    Dim i As Long: i = 0
    ReDim race_lists_(0)
    
    For i = 0 To UBound(fetched_items())
        If fetched_items(i) = "ŠJÃ“ú’ö" Then
            Do
                If InStr(fetched_items(i + UBound(race_lists_()) + 1), "/") = 0 Then
                    ReDim Preserve race_lists_(UBound(race_lists_()) - 1)
                    Exit Do
                End If
                race_lists_(UBound(race_lists_())) = fetched_items(i + UBound(race_lists_()) + 1)
                ReDim Preserve race_lists_(UBound(race_lists_()) + 1)
            Loop
            Exit For
        End If
    Next

    createRaceLists = race_lists_()
    
End Function

Private Function createRaceListParameters(ByRef fetched_items() As String) As Long
    
    If UBound(fetched_items()) = 0 Then
        createRaceListParameters = -1
    End If
    
    race_lists_() = createRaceLists(fetched_items())
    ReDim race_list_parameters_(UBound(race_lists_()))
    
    Dim i As Long
    For i = 0 To UBound(race_lists_())
        race_list_parameters_(i) = convertRaceListParameter(race_lists_(i))
    Next
    
    createRaceListParameters = 0

End Function

Public Function createCurrentRaceListParameters(ByRef fetched_items() As String) As String()

    If createRaceListParameters(fetched_items()) = -1 Then
    End If
    
    ReDim current_race_list_parameters_(0)
    Dim i As Long
    For i = 0 To UBound(race_lists_())
        If Left(race_list_parameters_(i), 1) = "c" Then
            current_race_list_parameters_(UBound(current_race_list_parameters_())) = race_list_parameters_(i)
            ReDim Preserve current_race_list_parameters_(UBound(current_race_list_parameters_()) + 1)
        End If
    Next
    
    If UBound(current_race_list_parameters_()) > 0 Then
        ReDim Preserve current_race_list_parameters_(UBound(current_race_list_parameters_()) - 1)
    End If
    
    createCurrentRaceListParameters = current_race_list_parameters_()
    
End Function

Private Function convertRaceListParameter(ByRef race_list As String) As String

     Dim pre_conversion_parameter As String
     pre_conversion_parameter = Left(race_list, 5)
     
     If DateDiff("d", Date, CDate(pre_conversion_parameter)) > 7 Then
        pre_conversion_parameter = "n" & pre_conversion_parameter
     ElseIf DateDiff("d", Date, CDate(pre_conversion_parameter)) <= 7 And DateDiff("d", Date, CDate(pre_conversion_parameter)) > 0 Then
        pre_conversion_parameter = "c" & pre_conversion_parameter
     ElseIf DateDiff("d", Date, CDate(pre_conversion_parameter)) < 0 And DateDiff("d", Date, CDate(pre_conversion_parameter)) > -7 Then
        pre_conversion_parameter = "p" & pre_conversion_parameter
     Else
        pre_conversion_parameter = "n" & pre_conversion_parameter
     End If
     
     convertRaceListParameter = Left(Replace(pre_conversion_parameter, "/", ""), 5)
     
End Function

Public Function createRaceDates(ByRef fetched_items() As String) As String()
    
    race_dates_() = Filter(fetched_items(), "“ú–Ú")
    createRaceDates = race_dates_()
    
End Function

Public Function createRaceDatesParameters() As Long
    On Error GoTo ErrorHandler
    
    ReDim race_date_parameters_(UBound(race_dates_()), LAST_RACE_NUMBER_)
    
    Const LAST_RACE_NUMBER_ As Integer = 12
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(race_dates_())
        For j = 0 To LAST_RACE_NUMBER_ - 1
            race_date_parameters_(i, j) = (Year(Now) & convertRaceDate(race_dates_(i)) & Format(j + 1, "00"))
        Next
    Next
   
    createRaceDatesParameters = 0
    Exit Function
ErrorHandler:
    Debug.Print Err.number & ":" & Err.Description, vbCritical & vbOKOnly, "ƒGƒ‰["
    createRaceDatesParameters = 1
End Function

Private Function convertRaceDate(ByVal race_date As String) As String
    
    Dim times As String
    times = Left(race_date, InStr(race_date, "‰ñ") - 1)
    
    Dim formatted_times As String
    formatted_times = Format(times, "00")

    Dim venue As String
    Dim event_date As String
    Select Case True
        Case race_date Like "*D–y*"
            venue = "01"
            event_date = Replace(Right(race_date, InStr(race_date, "D–y")), "“ú–Ú", "")
        Case race_date Like "*”ŸŠÙ*"
            venue = "02"
            event_date = Replace(Right(race_date, InStr(race_date, "”ŸŠÙ")), "“ú–Ú", "")
        Case race_date Like "*•Ÿ“‡*"
            venue = "03"
            event_date = Replace(Right(race_date, InStr(race_date, "•Ÿ“‡")), "“ú–Ú", "")
        Case race_date Like "*VŠƒ*"
            venue = "04"
            event_date = Replace(Right(race_date, InStr(race_date, "VŠƒ")), "“ú–Ú", "")
        Case race_date Like "*“Œ‹*"
            venue = "05"
            event_date = Replace(Right(race_date, InStr(race_date, "“Œ‹")), "“ú–Ú", "")
        Case race_date Like "*’†R*"
            venue = "06"
            event_date = Replace(Right(race_date, InStr(race_date, "’†R")), "“ú–Ú", "")
        Case race_date Like "*’†‹*"
            venue = "07"
            event_date = Replace(Right(race_date, InStr(race_date, "’†‹")), "“ú–Ú", "")
        Case race_date Like "*‹“s*"
            venue = "08"
            event_date = Replace(Right(race_date, InStr(race_date, "‹“s")), "“ú–Ú", "")
        Case race_date Like "*ã_*"
            venue = "09"
            event_date = Replace(Right(race_date, InStr(race_date, "ã_")), "“ú–Ú", "")
        Case race_date Like "*¬‘q*"
            venue = "10"
            event_date = Replace(Right(race_date, InStr(race_date, "¬‘q")), "“ú–Ú", "")
        Case Else
            Err.Raise number:=513, Description:="ƒŒ[ƒXî•ñ‚Ì•ÏŠ·‚ÅƒGƒ‰[‚ª”­¶‚µ‚Ü‚µ‚½"
    End Select
    
    Dim formatted_event_date As String
    formatted_event_date = Format(event_date, "00")

    convertRaceDate = venue & formatted_times & formatted_event_date

End Function

Public Function createCurrentRaceNames() As String

    


End Function

Public Function fetchRacePredictions(url_parameter_predictions)

    url_parameter_predictions_ = url_parameter_predictions

    Dim i As Integer
    For i = 0 To race_dates_.counts - 1
        Debug.Print url_parameter_predictions_(i)
        fetchRacePrediction (url_parameter_predictions_(i))
    Next

End Function

