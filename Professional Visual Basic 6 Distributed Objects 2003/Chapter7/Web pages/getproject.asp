<HTML>
<% 
Set objClient = Server.CreateObject("TaskObjects.Client")
objClient.SetSuperState Session("ClientState")
%>
<HEAD>
<TITLE>Client <% =objClient.Name %></TITLE>
</HEAD>
<BODY>
Client name:
<% =objClient.Name %><P>
<% 
  lngProject = Request.QueryString.Item(1)
  For Each objTemp In objClient.Projects
    if objTemp.ID = CLng(lngProject) then
      Set objProject = objTemp
      exit for
    end if
  next
%>
Project name:
<% = objProject.Name %><P>

<HR><P>
<TABLE border=1 cellPadding=1 cellSpacing=1 width=75%>
  <TR>
     <TD><STRONG>Task name</STRONG></TD>
     <TD><STRONG>Projected days</STRONG></TD>
     <TD><STRONG>Percent complete</STRONG></TD>
 </TR>

  <% For Each objTask In objProject.Tasks %>
    <TR>
      <TD><% =objTask.Name %></TD>
      <TD><% =objTask.ProjectedDays %></TD>
      <TD><% =objTask.PercentComplete %></TD>
 </TR>
  <% Next %>
</TABLE></P><BR>
</BODY>
</HTML>
