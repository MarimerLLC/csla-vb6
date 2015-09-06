<%
Set myVideo = Server.CreateObject("VideoObjects.Video")
myVideo.Load CLng(Request.QueryString("id"))
%>

<HTML>
<HEAD>
<TITLE><% Response.Write myVideo.Title %></TITLE>
</HEAD>
<BODY>
<P>
<%

Response.Write "</P>"
Response.Write "<P>" & myVideo.Studio & " Presents "
Response.Write Chr(34) & myVideo.Title & Chr(34) & "</P>"

strRating = Trim(myVideo.Rating)
Response.Write "<P>This movie is rated '" & strRating & "', "
Select Case strRating
Case "G"
  Response.Write "suitable for general audiences"
Case "PG"
  Response.Write "parental guidance suggested"
Case "PG-13"
  Response.Write "not suitable for children under age 13"
Case "R"
  Response.Write "children under 17 not admitted without parent"
Case "NR"
  Response.Write "not rated - for mature audiences"
End Select
Response.Write "</P>"

intAvail = myVideo.Tapes.Count

Response.Write "<P>We have " 
If intAvail = 0 Then
  Response.Write "no tapes "
ElseIf intAvail = 1 Then
  Response.Write "1 tape "
Else
  Response.Write intAvail & " tapes "
End If
Response.Write "available for rental</P>"

Set myVideo = Nothing
%>
</BODY>
</HTML>
