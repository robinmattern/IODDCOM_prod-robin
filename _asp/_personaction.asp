<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">
<title>Person Action Page</title>
</head>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->

<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<!--#config timefmt="%H:%M:%S"-->
<%
'=============================================
'Check for Session ID, if different, go back to default.asp
If trim(session("sessionid")) <> trim(session.sessionid) then
	Response.redirect "default.asp"
End If	


'=============================================
' Begin Processing Here
'=============================================
' Get Info from Calling Page
Session("now") = Now()
Session("CurrentDate") = FormatDateTime(Now(),2)
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Session("LastFrom") = Session("From")
from = ucase(trim(Request.QueryString("From")))
	'response.write "From = " & from
	'response.end
Session("From")= From
btn = ucase(trim(Request("btn")))
'response.write "From = " & From & "<br>"
'response.write "Button = " & btn & "<br>"
'response.end
'response.write "QueryString:  " & request.querystring
'response.end

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Decide what to do based on the calling page , From
'====================================================================================
Select Case FROM
case "_PERSON"

	select case btn
	case "SAVE CHANGES"	
		UpdatePerson()
		If Session("DeleteThisPerson") = "Yes" Then
			response.redirect "_personlist.asp"
		Else
			response.redirect "_person.asp?PersonID="&Session("PersonID")
		End If
		
	case "ADD A NEW USER"
		'response.write "ADD"
		'First Update
		UpdatePerson()

		'Then Add a New Record
		CreatedAt = Now()
		Organization = "-New"
		SQLStr = "INSERT tPerson (CustCode, ActiveUser, Organization, Logon, Password, CreatedAt, CreatedBy, PasswordChangedAt) VALUES ('" & Trim(session("mtstype")) & "', 'Yes', '" & Organization & "','-Email Address','newuser','" & CreatedAt & "', '" & session("ChangedBy") & "','" & CreatedAt & "')"
'		SQLStr = "INSERT tPerson (CustCode, ActiveUser, Organization, Logon, Password, CreatedAt, CreatedBy) VALUES ('" & Trim(session("mtstype")) & "', 'Yes', '" & Organization & "','-Email Address','newuser','" & CreatedAt & "', '" & session("ChangedBy") & "')"
		Conn.Execute(sqlStr)

		' Get PersonID
		SQLStr = "SELECT PersonID from tPerson where CreatedAt = '" & createdAt & "' and CreatedBy = '" & Session("ChangedBy") & "'" 
		Set rs = Conn.Execute(sqlStr)
		PersonID = rs("PersonID")

		response.redirect "_person.asp?PersonID="&PersonID

	case "USERS LIST"	

		UpdatePerson()
		response.redirect "_personlist.asp"

	End Select

'-----------------------------------------------------

case "_PERSONLIST"
	select case btn

	case "ADD A NEW USER"

		CreatedAt = Now()
		Organization = "-New"
		SQLStr = "INSERT tPerson (CustCode, ActiveUser, Organization, Logon, Password, CreatedAt, CreatedBy, PasswordChangedAt) VALUES ('" & Trim(session("mtstype")) & "', 'Yes', '" & Organization & "','-Email Address','newuser','" & CreatedAt & "', '" & session("ChangedBy") & "','" & CreatedAt & "')"
'		SQLStr = "INSERT tPerson (CustCode, ActiveUser, Organization, Logon, Password, CreatedAt, CreatedBy) VALUES ('" & Trim(session("mtstype")) & "', 'Yes', '" & Organization & "' , '-Email Address' ,'newuser','" & CreatedAt & "', '" & session("ChangedBy") & "')"
		'response.write SQLStr
		'response.end
		Conn.Execute(sqlStr)

		' Get PersonID
		SQLStr = "SELECT PersonID from tPerson where CreatedAt = '" & createdAt & "' and CreatedBy = '" & Session("ChangedBy") & "'" 
		Set rs = Conn.Execute(sqlStr)
		PersonID = rs("PersonID")

		response.redirect "_person.asp?PersonID="&PersonID

	End Select

End Select
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'FUNCTIONS

Function ConvertIN(str2format)
str2format = Trim(str2format)
'response.write "_incoming="&str2format&"<br>"&"Length="&Len(str2format)

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
Function UpdatePerson()
		totalrowcount = Cint(request("TotalRowCount"))
		currentrowcount = totalrowcount

		For i = 1 to TotalRowCount
			SetStr = ""   'Update Check
			If request(i&"LastName") <> request(i&"LastNameH") Then SetStr= SetStr+ "LastName = '" & replace(request(i&"LastName"),"'","''") & "', "
			If request(i&"FirstName") <> request(i&"FirstNameH") Then SetStr= SetStr+ "FirstName = '" & replace(request(i&"FirstName"),"'","''") & "', "
			If request(i&"Organization") <> request(i&"OrganizationH") Then SetStr= SetStr+ "Organization = '" & UCase(replace(request(i&"Organization"),"'","''")) & "', "
			If request(i&"Role") <> request(i&"RoleH") Then SetStr= SetStr+ "Role = '" & replace(request(i&"Role"),"'","''") & "', "
			If request(i&"Logon") <> request(i&"LogonH") Then SetStr= SetStr+ "Logon = '" & replace(request(i&"Logon"),"'","''") & "', "
			If request(i&"Password") <> request(i&"PasswordH") Then SetStr= SetStr+ "Password = '" & replace(request(i&"Password"),"'","''") & "', "
			If request(i&"Phone") <> request(i&"PhoneH") Then SetStr= SetStr+ "Phone = '" & replace(request(i&"Phone"),"'","''") & "', "				
			If request(i&"ActiveUser") <> request(i&"ActiveUserH") Then SetStr= SetStr+ "ActiveUser = '" & replace(request(i&"ActiveUser"),"'","''") & "', "
			If request(i&"ckDelete") <> request(i&"ckDeleteH") Then SetStr= SetStr+ "ckDelete = '" & replace(request(i&"ckDelete"),"'","''") & "', "
				If request(i&"ckDelete") = "X" then
					Session("DeleteThisPerson") = "Yes"
				Else
					Session("DeleteThisPerson") = "No"
				End If 
			If SetStr <> "" then
				SQLStr = "UPDATE tPerson Set " & SetStr & " ChangedAt = '" & Session("ChangedAt") & "', ChangedBy = '" & Session("ChangedBy") & "' WHERE PersonID = " & Request(i&"PersonIDH") 
				'response.write "SQL1: " & SQLStr & "<BR>"
				'response.end
				Conn.Execute(sqlStr)
			End if
		'Response.end

		'Delete the Users records that still have "-Email Address" in the logon field
		If request(i&"Logon") = "-Email Address" OR request(i&"Organiaztion") = "-New" then
			SQLStr = "DELETE tPerson FROM tPerson WHERE Logon = '-Email Address'"
			Conn.Execute(SQLStr)
			'response.write SQLStr
			'response.end
			response.redirect "_personList.asp"
		End If
		

		Next
		'response.write currentrowcount

		
		'Delete the Users records that were checked
		If Session("DeleteThisPerson") = "Yes" Then
			SQLStr = "DELETE tPerson FROM tPerson WHERE ckDelete = 'X'"
			Conn.Execute(SQLStr)
		End If
End Function

'=============================================


'END OF FUNCTIONS
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

%>
