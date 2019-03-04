VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SiteHanlder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IE As InternetExplorer

'/**
'* Create IE Object and set the destination URL
'*/
Private Sub Class_Initialize()

    Set IE = CreateObject("Internetexplorer.Application")
    IE.Visible = True
    
'    Dim today As String: today = Format(Date, "mmdd")
'    Dim siteURL As String: siteURL = "http://race.netkeiba.com/?pid=race_list&id=p" & today

    Dim siteURL As String: siteURL = "http://race.netkeiba.com/?pid=race_list&id=c0408"
    IE.navigate siteURL
    
    Do While IE.Busy = True Or IE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

End Sub

'/**
'* @param classNameOnSite
'* @retrun siteData
'*/

Private Function fetchSiteDataByClassName(classNameOnSite As String)

    Dim htmlDoc As HTMLDocument: Set htmlDoc = IE.document
    Dim elements As Object: Set elements = htmlDoc.getElementsByClassName(classNameOnSite)
    
    Dim element As Variant
    Dim preformattedText As String
    
    For Each element In elements
        preformattedText = preformattedText & Replace(element.innerText, vbCrLf, "") & ","
    Next element
    
    Dim siteData() As String
    siteData = Split(preformattedText, ",")
    ReDim Preserve siteData(UBound(siteData) - 1)
    
    fetchSiteDataByClassName = siteData

End Function

'/**
'*
'*/
Public Function defineFileName() As Variant

'     fetchSiteDataByClassName ("DateList_Box")

'    Dim elemets As Object
'    Set elements = fetchSiteDataByClassName("DataList_Box")
'
'    Dim rawText As String
'    Dim element As Variant
'
'    For Each element In elements
'        rawText = Replace(element.innerText, "/", "", 1)
'    Next element
'
'    Dim trimedText As String
'    trimedText = Replace(rawText, vbCrLf, "")
'
'    Dim separetedText As String
'    separetedText = Replace(trimedText, ")", "),")
'
'    Dim race_list_header() As String
'    race_list_header = Split(separetedText, ",")
'    ReDim Preserve race_list_header(UBound(race_list_header) - 1)
'
'    fetch_race_list_header = race_list_header

     fetchSiteDataByClassName ("kaisaidata")

End Function