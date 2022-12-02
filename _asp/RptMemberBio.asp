<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="inccreateconnection.asp"-->

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">

<title>Institute of Database Developers - Member List</title>
<meta name="Microsoft Theme" content="none, default">
</head>
<!--#INCLUDE FILE="incheader.asp"-->
<%
SQLStr = "SELECT FirstName + ' ' + LastName as MemberName, Bio, Skills FROM tMember Where Active = 'Y' ORDER BY lastname,Firstname"
set rs = conn.Execute(SQLStr)
%>

<body>


<div align="left">
<table width="85%" border="0">

<% do 
	IF rs.EOF THEN
		Exit Do
	End If
	%>	 
  <tr>
    <td width="20%" valign="top"><strong><%=rs("MemberName")%></strong></td>
    <td valign="top"><small><%=rs("Bio")%></small></td>
  </tr>
  <tr>
	<td> &nbsp; </td>
  </tr>
 	<%rs.MoveNext%>
<%loop%>
	</table>
</div>

</body>
</html>