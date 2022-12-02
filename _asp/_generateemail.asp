<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<!--#INCLUDE FILE="_incsessioncheck.asp"-->
<% Response.Buffer = True %>
<%Server.ScriptTimeOut =36000%>
<%
' Check Request Objects
If len(trim(Session("EmailFrom"))) = 0 then  ' No Sender
	Session("Message") = "No Sender. Cannot send email"
	response.redirect "_message.asp"
End If
If len(trim(request("To"))) > 0 then
	Session("EmailTo") = request("To")
End If	
If len(trim(request("Subject"))) > 0 then
	Session("EmailSubject") = request("Subject")
End If	
If len(trim(request("Message"))) > 0 then
	Session("EmailMessage") = request("Message")
End If	
%>
<!--#INCLUDE FILE="_incemail.asp"-->
<%
SendEMail Session("EmailFrom"),Trim(Session("EmailTo")),Session("EmailSubject"),Session("EmailMessage"),""
%>