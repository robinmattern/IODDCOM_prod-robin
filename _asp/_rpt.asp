<%@ LANGUAGE="vbscript" %>
<%
Session("ReportName") ="Sample Report"
Session("SQLStrReport") = "SELECT * from tperson"	
Session("SQLStrWhere") = ""	
'response.write session("SQLStr")
'response.end	
Response.Redirect "_GenerateReport.asp"
%>