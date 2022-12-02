<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="_incstrconvert.asp"-->
<!--#INCLUDE FILE="inccreateconnection.asp"-->

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">

<title>Institute of Database Developers - Member Contact</title>
<meta name="Microsoft Theme" content="none, default">
</head>

<!--#INCLUDE FILE="incheader.asp"-->
<body>
<!--#INCLUDE FILE="incbodyline.asp"-->

<%
SQLStr = "SELECT FirstName + ' ' + LastName as MemberName, * FROM tMember Where Active = 'Y' ORDER BY lastname,Firstname"
set rs = conn.Execute(SQLStr)
%>

<left>

<table cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" cellspacing="10">
<%do
		if rs.EOF then
			exit do
		end if%>

  <tr>
    <td><strong><a href="http://fido.gov/syschangepassword.asp?username=<%=rs("firstname") & "." & rs("lastname")%>"><%=rs("MemberName")%></a></strong></td>
    <td><small><a href="mailto:<%=strconvert(rs("Email"),"html")%>">Email Address</a></td>
    <td><small><%=rs("Phone1")%>&nbsp;&nbsp;&nbsp;<%=rs("Phone2")%></small></td>
  </tr>
<% rs.movenext
	loop%>
</table>
</left>
</body>
</html>