<%
If Request.Form("hdnAction") = Empty Then
%>
<HTML>
<HEAD>
<TITLE>Get Video Criteria</TITLE>
</HEAD>
<BODY>
<FORM ACTION="getvideos.asp" METHOD="POST">
      <INPUT TYPE=HIDDEN NAME="hdnAction" VALUE="0">
 Title:<BR> 
<INPUT TYPE=TEXT NAME="txtTitle"><P>
  Studio:<BR>
  <INPUT TYPE=TEXT NAME="txtStudio"><P>
  <INPUT TYPE=SUBMIT VALUE="Enter"> 
  <INPUT TYPE=RESET VALUE="Cancel">
</FORM>
</BODY>
</HTML>
<% 
Else
%>
  <%
  Set myVideos = Server.CreateObject("VideoObjects.Videos")
   strTitle = Request.Form("txtTitle")
  strStudio = Request.Form("txtStudio")
  myVideos.Load CStr(strTitle), CStr(strStudio)
  %>
  <HTML>
  <HEAD>
  <TITLE>Video List</TITLE>
  </HEAD>
  <BODY>
  <P>Here is a list of video titles from our wide selection:</P>
  <TABLE BORDER=1>
  <TR>
  <TD>Title</TD>
  <TD>Release date</TD>
  </TR>
  <%
  For Each Video In myVideos
    Response.Write "<TR>"
    %>
    <TD><A href="getvideo.asp?id=<% =Video.VideoID %>">
      <%= Video.Title %></A></TD>
    <%
    Response.Write "<TD>" & Video.ReleaseDate & "</TD>"
    Response.Write "</TR>"
  Next
  Set myVideos = Nothing
  %>
  </TABLE>
  </BODY>
  </HTML>

<%
End If
%>
