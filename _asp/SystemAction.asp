<%@ LANGUAGE="vbscript" %>
<% Response.Buffer = True %>
<% Response.Clear %>
<!--#INCLUDE FILE="incexpires.asp"-->
<!--#INCLUDE FILE="_incemail.asp"-->

<html>


<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">
<title><!--#INCLUDE FILE="incTitle.asp"--> System Action Page</title>
</head>

<!--#INCLUDE FILE="incbodyline.asp"-->
<!--#INCLUDE FILE="incheader.asp"-->

<!--#INCLUDE FILE="inccreateconnection.asp"-->
<!--#config timefmt="%H:%M:%S"-->

<%
'=============================================
' Begin Processing Here
'=============================================
' Get Info from Calling Page
Session("now") = Now()-.20833
Session("LastFrom") = Session("From")
requesttype = ucase(trim(Request.QueryString("type")))
Session("requesttype")= requesttype
from = ucase(trim(Request.QueryString("From")))
Session("From")= From
btn = ucase(trim(Request("btn")))

'-------------------------------------------------------------
' Decide what to do based on Type
If len(trim(requesttype)) > 0 then
	Select Case REQUESTTYPE
	case "GURUMAIL"
		response.redirect "Email.asp"
	End Select
End If

'-------------------------------------------------------------
' Decide what to do based on the calling page , From

Select Case FROM
'--------------------------------------------------------
case "EMAIL"
	' Check if sender is a member
	SQLStr = "SELECT * FROM tMember WHERE Email = '" & trim(request("emailfrom")) &"'"
	Response.write SQLStr & "<BR>"
	Set rs = conn.Execute(SQLStr)
	If rs.EOF then
		Mailto = ""
	Else	
		session("Useremail") = rs("email")
		session("username") = rs("firstname") & " " & rs("lastname")
		SQLStr = "SELECT Email,Lastname FROM tMember WHERE Active = 'Y' AND NOT Email IS NULL ORDER BY Email"
		'Response.write SQLStr
		Set rs = conn.Execute(SQLStr)
		Mailto = ""
		AddresseeLimit = 25
		AddresseeCount = 0
		TotalAddresseeCount = 0		
		Do while not rs.EOF
			TotalAddresseeCount = TotalAddresseeCount + 1		
			MailTo = MailTo & rs("email") & "; "
			AddresseeCount = AddresseeCount + 1
			If AddresseeCount = AddresseeLimit Then
				'response.write mailto
				SendEMail request("emailfrom"),MailTo,"IODD: " & Request("Subject"),Request("Message"),""
				MailTo = ""
				AddresseeCount=0
			End If	
			rs.Movenext
		Loop
		SendEMail request("emailfrom"),MailTo,"IODD: " & Request("Subject"),Request("Message"),""
	End if	
	rs.Close
	Set	 rs = Nothing

'---------------------------------------------------------
Case Else
	Session("Message") = "SystemAction Error - No Case Found " & Session("FROM")
	response.redirect "Message.asp"

End Select
' END OF PROCESSING
%>
<%
'=============================================
'Utility Functions
'=============================================
'=============================================

Function ConvertTimeToText(strval)
If IsNull(strval) OR strval = "1/1/1900" then
	ConvertTimeToText = ""
	Exit Function
End if	
strval = FormatDateTime(strval,3)
ap = "a"
If right(strval,2) ="PM" then
	ap = "p"
end if
If len(strval) = 10 then
	strval = "0" & strval
end if	
hr = Left(strval,2)
min = mid(strval,4,2)
'response.write "   " & strval & " - " & hr & " - " & min & " - " & ap
ConvertTimeToText = hr & min & ap
End Function 

'=============================================
Function ConvertTextToTime(strval)
If IsNull(strval) OR strval = "" then
	ConvertTextToTime = "" '#1/1/1900 12:00 PM# 'FormatDateTime("00:00",3)
	'response.redirect ("NullValue.asp")
	Exit Function
Else
hr = Left(strval,2)
min = mid(strval,3,2)
ap = right(strval,1)
	If ap = "p" then
		If hr <> "12" then
			hr = cstr(cint(hr)+12)
		End If
	End If		
ConvertTextToTime = hr & ":" & min
End If
End Function 

'=============================================

Function ConvertIN(str2format)
str2format = Trim(str2format)
'response.write "Incoming="&str2format&"<br>"&"Length="&Len(str2format)

If InStr(str2format, ",") = 0 Then
    ConvertIN = "('" & str2format & "')"
Else
    ConvertIN = "('"
    Do While InStr(str2format, ",") > 0
        NumSpaces = (InStr(str2format, ","))
        'response.write "NumSpaces= " &NumSpaces & "<br>"
        strpart = Trim(Left(str2format,NumSpaces-1))
        'response.write "strpart= "&strpart
        'response.write "<br>"&Len(str2format)& "-" &Len(strpart)&"<br>"&Len(str2format) - Len(strpart)-1
        str2format = Trim(Right(str2format,Len(str2format) - NumSpaces))
        ConvertIN = ConvertIN & strpart & "','"
    Loop
    ConvertIN = ConvertIN & Trim(str2format & "')")
End If
End Function

'=============================================
'END Utility Functions
'=============================================


'=============================================
'Log On Functions
'=============================================

Function checklogon()
CheckLogon = False %>
<!-- Perform Logon and Password check -->
<%If rs.EOF Then %>

<p align="center"><br>
<strong>Your Logon or PIN is not correct.</strong><br>
<br>
<strong>Please click the Return to Logon button.</strong> </p>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td align="center" width="100%" valign="middle"><form method="POST" Action="../default.asp">
      <p><font color="#400040"><input type="submit" value="Return to Logon" name="DataAction"></font></p>
    </form>
    </td>
  </tr>
</table>
<p>
<!--#INCLUDE FILE="incfooter.asp"-->

<% rs.Close
SET rs = Nothing%> 
<%Exit Function%> 
<% End If %> 

<!-- Perform Active User check --> 
<%If rs("ActiveUser")<>"Yes"  Then %> </p>

<p align="center"><br>
<strong>Your Account is not active. Please contact our office to have your user information
updated.</strong><br>
<br>
<strong>Please click the Return to Logon button.</strong> </p>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td align="center" width="100%" valign="middle"><form method="POST" Action="../default.asp">
      <p><font color="#400040"><input type="submit" value="Return to Logon" name="DataAction"></font></p>
    </form>
    </td>
  </tr>
</table>

<p>
<!--#INCLUDE FILE="incfooter.asp"-->
	<% rs.Close
	SET rs = Nothing%><%Exit Function%><% End If %><%

'Get Browser Info
Set bc = Server.CreateObject("MSWC.BrowserType")   
Session("Browser") = bc.browser & " " & bc.version
SESSION("UserID")=rs("UserID")
SESSION("UserEmail")=rs("Email")
SESSION("PermissionLevel")=rs("PermissionLevel")
SESSION("AgencyAbbr")=rs("AgencyAbbr")
Session("UserName") = rs("FirstName") & " " & rs("LastName")
Session("UserFirstName") = rs("FirstName")
Session("ChangedBy") = trim(rs("FirstName"))& " " & trim(rs("LastName"))
SESSION("Message")=""
SQLStr="Update tblUser SET LASTLOGONDATE = GETDATE(), BROWSER =  ' " & SESSION("BROWSER") & " ' WHERE UserID = " & SESSION("UserID")
SET rs = NT.Execute(SQLStr)
SET rs = Nothing
CheckLogon = True

End FUNCTION%></html>