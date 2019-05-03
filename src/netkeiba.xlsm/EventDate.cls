VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private items_() As String
Private event_dates_() As String     '04/13(“y), 04/14(“ú), 04/20(“y), 04/21(“ú),
Private event_dates_parameters_() As String       'p0420, c0421, n0422,
Private current_event_dates_()  As String   '04/20(“y), 04/21(“ú),
Private current_event_date_parameters_() As String       'c0420, c0421,

Private Sub Class_Initialize()
    
End Sub

Public Property Let items(ByRef items() As String)
    items_() = items()
End Property

Public Property Get event_dates() As String()
    event_dates = event_dates_()
End Property

Public Property Get current_event_dates() As String()
    current_event_dates = current_event_dates_()
End Property

Public Property Get current_event_dates_parameters() As String()
    current_event_dates_parameters = current_event_date_parameters_()
End Property

Public Function analyzeItems()
    
    If createRaceLists() <> 0 Then
    End If
    
    If createRaceListParameters() <> 0 Then
    End If
    
    If createCurrentRaceListParameters() <> 0 Then
    End If
    
End Function

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

Private Function createRaceLists() As Long

    Dim i As Long: i = 0
    ReDim event_dates_(0)
    
    For i = 0 To UBound(items_())
        If items_(i) = "ŠJÃ“ú’ö" Then
            Do
                If InStr(items_(i + UBound(event_dates_()) + 1), "/") = 0 Then
                    ReDim Preserve event_dates_(UBound(event_dates_()) - 1)
                    Exit Do
                End If
                event_dates_(UBound(event_dates_())) = items_(i + UBound(event_dates_()) + 1)
                ReDim Preserve event_dates_(UBound(event_dates_()) + 1)
            Loop
            Exit For
        End If
    Next
    
    createRaceLists = 0
    
End Function

Private Function createRaceListParameters() As Long
    
    ReDim event_dates_parameters_(UBound(event_dates_()))
    ReDim current_event_dates_(0)
    
    Dim i As Long
    For i = 0 To UBound(event_dates_())
        event_dates_parameters_(i) = convertRaceListParameter(event_dates_(i))
    Next
       
    If UBound(current_event_dates_()) > 0 Then
        ReDim Preserve current_event_dates_(UBound(current_event_dates_()) - 1)
    End If
    
    createRaceListParameters = 0

End Function

Private Function convertRaceListParameter(ByVal event_date As String) As String

     Dim pre_conversion_parameter As String
     pre_conversion_parameter = Left(event_date, 5)
     
     If DateDiff("d", Date, CDate(pre_conversion_parameter)) > 7 Then
        pre_conversion_parameter = "n" & pre_conversion_parameter
        
     ElseIf DateDiff("d", Date, CDate(pre_conversion_parameter)) <= 7 And DateDiff("d", Date, CDate(pre_conversion_parameter)) > 0 Then
        current_event_dates_(UBound(current_event_dates_())) = event_date
        ReDim Preserve current_event_dates_(UBound(current_event_dates_()) + 1)
        pre_conversion_parameter = "c" & pre_conversion_parameter
 
     ElseIf DateDiff("d", Date, CDate(pre_conversion_parameter)) < 0 And DateDiff("d", Date, CDate(pre_conversion_parameter)) > -7 Then
        pre_conversion_parameter = "p" & pre_conversion_parameter
        
     Else
        pre_conversion_parameter = "n" & pre_conversion_parameter
     End If
     
     convertRaceListParameter = Left(Replace(pre_conversion_parameter, "/", ""), 5)
     
End Function

Private Function createCurrentRaceListParameters() As Long

    ReDim current_event_date_parameters_(0)
    Dim i As Long
    For i = 0 To UBound(event_dates_())
        If Left(event_dates_parameters_(i), 1) = "c" Then
            current_event_date_parameters_(UBound(current_event_date_parameters_())) = event_dates_parameters_(i)
            ReDim Preserve current_event_date_parameters_(UBound(current_event_date_parameters_()) + 1)
        End If
    Next
    
    If UBound(current_event_date_parameters_()) > 0 Then
        ReDim Preserve current_event_date_parameters_(UBound(current_event_date_parameters_()) - 1)
    End If
    
    createCurrentRaceListParameters = 0
    
End Function
