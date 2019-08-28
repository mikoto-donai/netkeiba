VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Netkeiba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Enum RACE_NUMBER   'レース番号
    FIRST = 1
    LAST = 12
End Enum

Private login_id_ As String  'netkeiba アカウント  ex. hoge@fuga.com
Private login_pass_ As String  'netkeiba パスワード  ex. piyo

Private race_dates_() As String  'レース日  ex. race_date(0) / 2回東京7日目
Private race_parameters_() As String   '各レースのURLパラメータ  ex.  race_parameters_(0) / 2019050207
Private file_names_() As String  'セル値格納先のファイル名 ex. file_names(0,1) / 2回東京7日目:2R
Private race_results_() As Variant    'セル値  ex.  race_results_(0, 1) / 2回東京7日目:2R のセル値

Private Sub Class_Initialize()
    ReDim race_years_(0)
    ReDim race_places_(0)
End Sub

Public Property Let login_id(ByVal login_id As String)
    login_id_ = login_id
End Property

Public Property Let login_pass(ByVal login_pass As String)
    login_pass_ = login_pass
End Property

Public Property Let url(ByVal url As String)
    url_ = url
End Property

Public Property Get fileNames() As String()
    fileNames = file_names_()
End Property

Public Property Get raceResults() As Variant()
    raceResults = race_results_()
End Property

Public Function analyzeItems(ByVal race_year As Long, ByRef items As Variant)
        
    If loginNetkeiba() = -1 Then
        Err.Raise Number:=513, Description:="ログインに失敗しました"
        Exit Function
    End If
        
    If createRaceDates(items) = -1 Then
        Err.Raise Number:=513, Description:="レース日の作成に失敗しました"
        Exit Function
    End If
    
    If createRaceParameters(race_year) = -1 Then
        Err.Raise Number:=513, Description:="レースパラメータの作成に失敗しました"
        Exit Function
    End If
    
    If fetchRaceResults() = -1 Then
        Err.Raise Number:=513, Description:="各レースの取得に失敗’しました"
        Exit Function
    End If

End Function

Private Function loginNetkeiba() As Long

    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = False
    
    ie.Navigate "https://regist.netkeiba.com/account/?pid=login"
    
    Do While ie.Busy = True Or ie.ReadyState <> 4
      DoEvents
    Loop
    
    Dim DOM As Object
    For Each DOM In ie.document.getElementsByTagName("input")
      If InStr(DOM.Name, "login_id") > 0 Then
          DOM.Value = login_id_
      End If
      
      If InStr(DOM.Name, "pswd") > 0 Then
          DOM.Value = login_pass_
      End If
    Next
    
    For Each DOM In ie.document.getElementsByTagName("input")
      If InStr(DOM.Type, "image") > 0 Then
          DOM.Click
      End If
    Next
    
    loginNetkeiba = 0
 
End Function

Private Function createRaceDates(ByRef items As Variant) As Long

    Dim i As Long
    ReDim race_dates_(0)
    
    For i = LBound(items, 1) To UBound(items, 1)
        If InStr(items(i, 1), "回") > 0 Then
            race_dates_(UBound(race_dates_)) = Mid(items(i, 1), InStr(items(i, 1), " ") + 1) & "目"
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
        Case race_date Like "*札幌*"
            event_venue = "01"
            event_date = Mid(race_date, InStr(race_date, "札幌") + 2)
        Case race_date Like "*函館*"
            event_venue = "02"
            event_date = Mid(race_date, InStr(race_date, "函館") + 2)
        Case race_date Like "*福島*"
            event_venue = "03"
            event_date = Mid(race_date, InStr(race_date, "福島") + 2)
        Case race_date Like "*新潟*"
            event_venue = "04"
            event_date = Mid(race_date, InStr(race_date, "新潟") + 2)
        Case race_date Like "*東京*"
            event_venue = "05"
            event_date = Mid(race_date, InStr(race_date, "東京") + 2)
        Case race_date Like "*中山*"
            event_venue = "06"
            event_date = Mid(race_date, InStr(race_date, "中山") + 2)
        Case race_date Like "*中京*"
            event_venue = "07"
            event_date = Mid(race_date, InStr(race_date, "中京") + 2)
        Case race_date Like "*京都*"
            event_venue = "08"
            event_date = Mid(race_date, InStr(race_date, "京都") + 2)
        Case race_date Like "*阪神*"
            event_venue = "09"
            event_date = Mid(race_date, InStr(race_date, "阪神") + 2)
        Case race_date Like "*小倉*"
            event_venue = "10"
            event_date = Mid(race_date, InStr(race_date, "小倉") + 2)
        Case Else
            Err.Raise Number:=513, Description:="レース情報の変換でエラーが発生しました"
    End Select
 
    event_date = Format(Replace(event_date, "日目", ""), "00")

    convertRaceDateIntoParameter = event_venue & event_time & event_date
    
End Function

Private Function fetchRaceResults() As Long

    ReDim race_results_(UBound(race_parameters_()), RACE_NUMBER.LAST - 1)
    ReDim file_names_(UBound(race_parameters_()), RACE_NUMBER.LAST - 1)

    Const INSPECTION_CELL_ROW = 1
    Const INSPECTION_CELL_COLUMN = 1
    Const INSPECTION_CELL_VALUE = "着"

    Dim o_fetcher As New Fetcher
    Const targetSheetName = "fetched_data"
    o_fetcher.configureTargetSheet targetSheetName
    
    Dim i As Integer
    Dim j As Integer
    Dim v As Variant
    
    For i = LBound(race_parameters_()) To UBound(race_parameters_())
        For j = RACE_NUMBER.FIRST - 1 To RACE_NUMBER.LAST - 1

            o_fetcher.url = "https://db.netkeiba.com/race/" & race_parameters_(i) & Format(j + 1, "00")
            o_fetcher.fetchItems 1

            If o_fetcher.items(INSPECTION_CELL_ROW, INSPECTION_CELL_COLUMN) = INSPECTION_CELL_VALUE Then
                file_names_(i, j) = race_dates_(i) & ":" & j + 1 & "R"
                race_results_(i, j) = o_fetcher.items
            End If
            
        Next
        FileIO.writeLog race_dates_(i) & "を取得しました"

        'If i = 1 Then
        '    Exit For
        'End If
            
    Next
    
    fetchRaceResults = 0
    
End Function
