<%@ LANGUAGE="VBScript" %>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="incexpires.asp"-->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">

<title>debug.asp</title>
</head>
<body bgcolor="#FFFFFF"><font size="1">

<% 
On Error Resume Next
DBType = UCASE(Trim(Request.QueryString("Type")))
If DBType = "" then
	DBType = "ALL"
End IF

If DBType = "ALL" OR dbtype = "A" then
Response.Write("<P>APPLICATION CONTENTS VARIABLES:<br>")
Response.Write("----------------<br>")
For Each Key in Application.Contents
	If Left(Key,3) = "ARR" then    'We have an array
		response.write Key & ": "
		arrayname = Application.Contents(Key)
		for i = 0 to ubound(arrayname,2)
			response.write arrayname(0,i) & ", "
		next		
		Response.Write("<br>")
	Else
		Response.Write( Key & " = " & Application.Contents(Key) & "<br>")
	End If
Next
End IF

If DBType = "ALL" OR dbtype = "S" then
Response.Write("<P>SESSION CONTENTS VARIABLES:<br>")
Response.Write("----------------<br>")
For Each Key in Session.Contents
	If Instr(Key,"RUNTIME") = 0 then
		If Left(Key,3) = "ARR" then    'We have an array
			response.write Key & ": "
			arrayname = Session.Contents(Key)
			for i = 0 to ubound(arrayname,2)
				response.write arrayname(0,i) & ", "
			next		
			Response.Write("<br>")
		Else
			Response.Write( Key & " = " & Session.Contents(Key) & "<br>")
		End If
	End IF 
Next
End IF

If DBType = "ALL" OR dbtype = "F" then
Response.Write("<P>FORM VARIABLES:<br>")
Response.Write("----------------<br>")
For Each Key in Request.Form
Response.Write( Key & " = " & Request.Form(Key) & "<br>")
Next
End IF

If DBType = "ALL" OR dbtype = "Q" then
Response.Write("<P>QUERY STRING VARIABLES:<br>")
Response.Write("-----------------------<br>")
For Each Key in Request.QueryString
Response.Write( Key & " = " & Request.QueryString(Key) & "<br>")
Next
End IF

If DBType = "ALL" OR dbtype = "C" then
Response.Write("<P>COOKIE VARIABLES:<br>")
Response.Write("-----------------<br>")
For Each Key in Request.Cookies
Response.Write( Key & " = " & Request.Cookies(Key) & "<br>")
Next
End IF

If DBType = "ALL" OR dbtype = "C" then
Response.Write("<P>CLIENT CERTIFICATE VARIABLES:<br>")
Response.Write("-----------------<br>")
For Each Key in Request.ClientCertificate
Response.Write( Key & " = " & Request.ClientCertificate(Key) & "<br>")
Next
End IF

If DBType = "ALL" OR dbtype = "S" then
Response.Write("<P>SERVER VARIABLES:<br>")
Response.Write("-----------------<br>")
For Each Key in Request.ServerVariables
Response.Write( Key & " = " & Request.ServerVariables(Key) & "<br>")
Next

End If
%>
</font></body></html>


