<%@ LANGUAGE="vbscript" %>
<% Response.Buffer = True %>
<% Response.ExpiresAbsolute=#May 31,1996 13:30:15# %> 

<%Function LogonFailed()%>
<! Logon not found ->
<meta name="Microsoft Theme" content="none, default">
<body>
<table Color="White" border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td colspan="3" bgcolor="Red"><font size="5" face="Arial, Helvetica, sans-serif">&nbsp;Logon Error&nbsp; </font></td>
    </tr>
</table>
	<br>
	<strong>Your password is not correct.</strong><br><br>
	<strong>Please press the back button or arrow to re-enter.</strong></font>
	</p>

<%End FUNCTION%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">

<title>Logon Action</title>
</head>
<body>

<%
IF Ucase(Request("Password")) = "DATABASE" Then
	Session("CanUpdate")=2
	Response.Redirect "MemberList.asp"
END IF
IF Ucase(Request("Password")) = "GUEST" Then
	Session("CanUpdate")=1
	Response.Redirect "MemberList.asp"
END IF

Call LogonFailed()
%>

</body>
</html>


