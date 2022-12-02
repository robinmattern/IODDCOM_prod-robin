<%@ LANGUAGE="vbscript" %>
<%
Session("ReportName") = "Message Report"
Session("SQLStr") =  "SELECT MessageDate AS Date, MessageSubject As Subject, MessageBody AS Message FROM tMessage where MessageToPerson = '" & session("email") & "' ORDER BY MessageDate DESC"	
'response.write Session("SQLStr")
'response.end
Response.Redirect "GenerateReport.asp"
%>