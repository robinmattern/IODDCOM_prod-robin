<%@ LANGUAGE="VBScript" %>
<!--#INCLUDE FILE="_incsessioncheck.asp"-->
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
<%

session("ReportName") = "Users Report"
session("sqlstr") = "SELECT * from tperson where Organization = '" & Session("AgencyAbbr") & "' ORDER BY Organization, Lastname, Firstname"
response.redirect "_generatereport.asp"
%>

</body>
</html>