<!--#INCLUDE FILE="inccreateconnection.asp"-->
<%
' SendEmail
'test SendEmail "bruce.troutman@8020data.com","bruce.troutman@fido.gov;btroutma@8020data.com;","Test from SendASPMail","Test Message",""

'=========================================================================
Function SendEmail( strFrom,strTo,strSubject,strMessage,strNextPage)
'strNextPage = "" ' This is broken
'response.write "SendEmail is SendASPMail" & "<BR>"
'response.write strFrom & "<BR>"
'response.write strTo & "<BR>"
'response.write strSubject &  "<BR>"
'response.write strMessage & "<BR>"
'response.end
msg = ""
If "" & strFrom = "" then msg = msg & "From is empty. "
If "" & strTo = "" then msg = msg & "To is empty. "
If "" & strSubject = "" then msg = msg & "Subject is empty. "
If "" & strMessage = "" then msg = msg & "Message is empty. "
If msg <> "" then
	session("message") = msg
	response.redirect "_message.asp"
End If	

' Handle Apostrophe's
strFrom = replace(strFrom,"'","''")
strTo = replace(strTo,"'","''")
strSubject = replace(strSubject,"'","''")
strMessage = replace(strMessage,"'","''")
strToList = StrTo

' Store in tMessage for each recipient
' Breakout each recipient NOTE: separator must be ;
RecList = StrTo
Do while "" & RecList <> ""
	If instr(RecList,";") = 0 then  ' Only 1 recipient
      	Recipient = Trim(RecList)
      	Sqlstr = "Insert into tMessage (MessageFrom, MessageTolist,MessageToPerson,MessageSubject,MessageBody,MessageDate,CreatedAt,CreatedBy) Values ('" & strFrom & "','" & StrTo & "','" & Recipient & "','" & StrSubject & "','" & StrMessage & "','" & now() & "','" & now() & "','" &  strFrom  & "')"
      	Set rsmsg = Conn.Execute(CleanSQL(sqlstr))
      	Exit Do
	Else		
		Recipient = trim(left(RecList,instr(reclist,";")-1))
		Reclist = Mid(Reclist,instr(reclist,";")+1)
    	Sqlstr = "Insert into tMessage (MessageFrom, MessageTolist,MessageToPerson,MessageSubject,MessageBody,MessageDate,CreatedAt,CreatedBy) Values ('" & strFrom & "','" & strTo & "','" & Recipient & "','" & StrSubject & "','" & StrMessage & "','" & now() & "','" & now() & "','" &  strFrom  & "')"
	   	Set rsmsg = Conn.Execute(CleanSQL(sqlstr))
	end if
Loop 
Set rsmsg = Nothing


' Send Message
Set Mail = Server.CreateObject("Persits.MailSender")
  Mail.Host = "mail.fido.gov"     ' Specify a valid SMTP server
  Mail.Helo="mail.fido.gov"
  Mail.Charset="iso-8859-1"
Mail.From = strFrom	  ' Specify authorized account
Mail.MailFrom = strFrom	  ' Specify authorized account
Mail.FromName = strFrom	  ' Specify authorized account
If len(trim(strTo)) = 0 then ' No recipients
	session("message") = "No email recipients"
	response.redirect = "_message.asp"
End if	
' Create Individual addresses from list
sTo = strTo
if not instr(sTo,";") then
	sTo = sTo & ";"
end if	
do while Len(Trim(sTo))>0
	Address = Left(sTo,instr(sTo,";")-1)
	If len(trim(Address)) > 0 then 
		'response.write "Address: " & Address & "<br>"
		Mail.AddAddress Address, Address
	End If
	sTo = Mid(sTo,instr(sTo,";")+1)
loop
Mail.Subject = strSubject
Mail.Body = strMessage
'	On Error Resume Next
Mail.Send
Set Mail = Nothing
Session("Message") = ""
If Err <> 0 Then
	Session("Message") = "ASP Email Error encountered: " & Err.Number & " " & Err.Description & "<br>" & " fm:" & strFrom & "<br> to:" & strTo & "<br> subj:" & strSubject & "<br> message:" & strMessage & "<br> nextpage:" & strNextPage
	response.redirect "_message.asp"
Else 
	If len(trim(strNextPage)) > 0 then
		response.redirect strNextPage
	Else
		Session("Message") = "Email successfully sent to: <br><br>" & strToList & "."
		response.redirect "_message.asp"
	End If			
End if
'response.write "Session(Message): " & Session("Message") & "<br>"
'response.end
End Function




%>