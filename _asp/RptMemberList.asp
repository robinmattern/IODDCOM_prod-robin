<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="_incstrconvert.asp"-->

<!--#INCLUDE FILE="inccreateconnection.asp"-->

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">

<title>Institute of Database Developers - Member List</title>
<meta name="Microsoft Theme" content="none, default">
</head>
<!--#INCLUDE FILE="incheader.asp"-->
<body>
<!--#INCLUDE FILE="incbodyline.asp"-->
<%
SQLStr = "SELECT FirstName + ' ' + LastName as MemberName, * FROM tMember Where Active = 'Y' ORDER BY lastname,Firstname"
set rs = conn.Execute(SQLStr)
%>


<div align="left">

<table cellspacing="6" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
<%do
		if rs.EOF then
			exit do
		end if%>

  <tr>
    <td width=20%><strong><%=rs("MemberName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></td>
    <td width=40%><small><%=rs("Phone1")%>&nbsp;&nbsp;&nbsp;<%=rs("Phone2")%></small></td>
    <td><small><a href="mailto:<%=strconvert(rs("Email"),"html")%>">Email Address</a></td>
  </tr>
  <tr>
    <td><small>&nbsp;</small></td>
    <td><small><%=rs("Company")%></small></td>
    <td><small>Web: <a href="<%=rs("WebSite")%>"><%=rs("WebSite")%></a></small></td>
  </tr>
  <tr>	
		<td> &nbsp;</td>
		<td><small><%=rs("Address1")%></small></td>
  </tr>
  <tr>	
		<td> &nbsp;</td>
	  <td><small><%=rs("City") & ", " & rs("State") & " " & rs("Zip")%></small></td>
  </tr>
  <tr>	
		<td> &nbsp;</td>
	  <td><small><%=rs("bio")%></small></td>
  </tr>
  <tr>	
	  <td>&nbsp;</td>
  </tr>
<% rs.movenext
	loop%>
</table>

</div>
</body>
</html>
