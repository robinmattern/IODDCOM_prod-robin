<%@ LANGUAGE="VBSCRIPT" %>
<%
  intID = Request.QueryString ("ID")
  If intID > 0 Then 'Value obtained.
    SQLQuery = "SELECT DISTINCTROW Message.ID, Message.Subject, " _
             & "Message.From, Message.Email, Message.Body, Message.When, " _
             & "Message.MsgLevel FROM Message " _
             & "WHERE (((Message.ID)= " & intID & "));"
    Set RS= conn.Execute(SQLQuery)
  End If
%>
<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">

<title><% = RS.Fields("Subject") %></title>
<meta name="Microsoft Theme" content="none, default">
</head>

<body>
<a HREF="../forum.asp">

<p align="center"><img BORDER="0" SRC="../BackToTheForum.gif" WIDTH="221" HEIGHT="17"></a> </p>

<hr>

<table WIDTH="640" BORDER="0" CELLPADDING="0" CELLSPACING="0" VALIGN="TOP">
  <tr>
    <td ALIGN="RIGHT" VALIGN="TOP" WIDTH="95"><b>Subject:</b><br>
    <b>From:</b><br>
    <b>Host:</b><br>
    <b>Date:</b></td>
    <td VALIGN="TOP" WIDTH="12"></td>
    <td VALIGN="TOP" WIDTH="533"><b><% = RS.Fields("Subject") %></b><br>
    <b><% = RS.Fields("From") %></b><br>
    <b><%
strHostName = Request("REMOTE_HOST")
strUserName = Request("REMOTE_USER")
If Len(strUserName) Then strHostName = strUsername & "@" & strHostName
Response.Write(strHostName) %></b><br>
    <b><% = RS.Fields("When") %></b></td>
  </tr>
</table>

<table VALIGN="TOP">
  <tr>
    <td><blockquote>
<%= RS.Fields("Body") %>
    </blockquote>
    </td>
  </tr>
</table>

<hr>

<p><br>
<!-- Now get a list of all apropriate messages --></p>

<p align="center"><a HREF="PostForm.asp?ID=<% = intID %>"><img BORDER="0" SRC="PostReply.gif" WIDTH="122" HEIGHT="22"></a> </p>
</body>
</html>
