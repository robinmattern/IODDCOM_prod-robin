<%@ LANGUAGE="VBScript" %>
<% Response.Buffer = True %>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Points of Contact (POC)</title>
</head>

<body>

<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<%' <!--#INCLUDE FILE="dminccreateconnection.asp"--> %>
<%
session("ReportName") = "Points of Contact (POC) List"
session("sqlstrwhere") = ""
session("sqlstr") = "SELECT Organization as Agency, FirstName + ' ' + Lastname as Name , Phone, Logon as Email, LastLogonDate as LastLogon from tperson where Role = 'POC' and ActiveUser = 'Yes' ORDER BY Organization, Lastname, Firstname"
response.redirect "generatereport.asp"
%>

</body>
</html>