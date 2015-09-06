VERSION 5.00
Begin {17016CEE-E118-11D0-94B8-00A0C91110ED} GetVideos 
   ClientHeight    =   5445
   ClientLeft      =   750
   ClientTop       =   1425
   ClientWidth     =   7320
   _ExtentX        =   12912
   _ExtentY        =   9604
   MajorVersion    =   0
   MinorVersion    =   8
   StateManagementType=   1
   ASPFileName     =   ""
   DIID_WebClass   =   "{12CBA1F6-9056-11D1-8544-00A024A55AB0}"
   DIID_WebClassEvents=   "{12CBA1F5-9056-11D1-8544-00A024A55AB0}"
   TypeInfoCookie  =   15
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   3
      BeginProperty WebItem1 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "DisplayVideo"
         DISPID          =   1280
         Template        =   "displayvideo.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{0672650A-56C1-11D2-93E5-00104B4C8457}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "C:\displayvideo.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem2 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "ListVideos"
         DISPID          =   1281
         Template        =   "listvideos.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{8ECCB7D4-56B9-11D2-93E5-00104B4C8457}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "C:\listvideos.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem3 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "SearchForm"
         DISPID          =   1282
         Template        =   "searchform1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{067264EC-56C1-11D2-93E5-00104B4C8457}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "C:\InetPub\wwwroot\wroxvideo\searchform.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
   EndProperty
   NameInURL       =   "getvideos"
End
Attribute VB_Name = "GetVideos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objVideos As Videos
Private objVideo As Video

Private Sub WebClass_Start()
  With Request.Form
     If .Count = 0 Then

       SearchForm.WriteTemplate
     Else
       Set objVideos = New Videos
       objVideos.Load .Item("txtTitle"), .Item("txtStudio")
       ListVideos.WriteTemplate
     End If
  End With
End Sub

Private Sub ListVideos_ProcessTag(ByVal TagName As String, _
    TagContents As String, SendTags As Boolean)
  Dim strResponse As String
  Dim objItem As VideoDisplay
  
  For Each objItem In objVideos
    strResponse = strResponse & "<TR>"
    
     strResponse = strResponse & "<TD><A href=""" & _
      URLFor(DisplayVideo, Format$(objItem.VideoID)) & """>" & _
      objItem.Title & "</A></TD>"
   
    strResponse = strResponse & "<TD>" & objItem.ReleaseDate
    strResponse = strResponse & "</TD>"

    strResponse = strResponse & "</TR>"
  Next
  Set objVideos = Nothing
  TagContents = strResponse
End Sub

Private Sub DisplayVideo_UserEvent(ByVal EventName As String)
  Set objVideo = New Video
  objVideo.Load Val(EventName)
  DisplayVideo.WriteTemplate
End Sub

Private Sub DisplayVideo_ProcessTag(ByVal TagName As String, _
    TagContents As String, SendTags As Boolean)
  Dim intAvail As Integer

  Select Case TagContents
  Case "Title"
    TagContents = objVideo.Title
  Case "Studio"
    TagContents = objVideo.Studio
  Case "Rating"
    TagContents = objVideo.Rating
  Case "RatingText"
    Select Case objVideo.Rating
    Case "G"
      TagContents = "suitable for general audiences"
    Case "PG"
      TagContents = "parental guidance suggested"
    Case "PG-13"
      TagContents = "not suitable for children under age 13"
    Case "R"
      TagContents = "children under 17 not admitted without parent"
    Case "NR"
      TagContents = "not rated - for mature audiences"
    End Select
  Case "TapeCount"
    intAvail = objVideo.Tapes.Count
    If intAvail = 0 Then
      TagContents = "no tapes "
    ElseIf intAvail = 1 Then
      TagContents = "1 tape "
    Else
      TagContents = intAvail & " tapes "
    End If
  End Select
End Sub

