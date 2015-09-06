<HTML>
<% 
Set objClient = Server.CreateObject("TaskObjects.Client")
lngID = Request.Form("txtID")
on error resume next
objClient.Load CLng(lngID)

if err = 3021 then
%>
   <HEAD>
   <TITLE>Client not found</TITLE>
   </HEAD>
   <BODY>
   Client not found in the database.
   </BODY>

<%   elseif err <> 0 then 
%>
   <HEAD>
   <TITLE>Client not found</TITLE>
   </HEAD>
   <BODY>
   An error has occured while retrieving the client<br>
   The error number is <% =err %><br>
   </BODY>

<% else
  Session("ClientState") = objClient.GetSuperState
%>
  <HEAD>
  <TITLE>Client <% =objClient.Name %></TITLE>
  </HEAD>
  <BODY>
  Client ID:
  <% =objClient.ID %><P>
  Client name:
  <% =objClient.Name %><P>
  Client phone:
  <% =objClient.Phone %><P>

  <HR><P>
  <TABLE border=1 cellPadding=1 cellSpacing=1 width=75%>
    <TR>
      <TD><STRONG>Project name</STRONG></TD>
    </TR>

    <% For Each objProject In objClient.Projects %>
      <TR>
        <TD><A href="getproject.asp?project=<%=objProject.ID%>">
          <% =objProject.Name %></A>
        </TD>
      </TR>
    <% Next %>
  </TABLE></P><BR>
  </BODY>
<% end if %>
</HTML>


