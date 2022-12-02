<%@ LANGUAGE="VBScript" %>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="_incsessioncheck.asp"-->
<!--#INCLUDE FILE="_incexpires.asp"-->
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
%>
	<center>
	<table border="2" width="90%" cellspacing="0" cellpadding="0">
	<tr>
	<td colspan="2" align="left" bgcolor="#99CCFF">
      <p align="center"><font color="#000080" face="Arial"><b>Debug Page    (<%=Now()%>)</b></font>
    </td>
	</tr>
	<tr>
	<td colspan="2" align="left" bgcolor="#99CCFF">
      <p align="center"><font color="#000080" size="1" face="Arial"><b>APPLICATION CONTENTS VARIABLES</b></font></p>
    </td>
	</tr>
<%
For Each Key in Application.Contents
	If ucase(Left(Key,7)) = "LOOKUP_" then    'We have an array 
		%>
		<tr>
		<td width="25%" valign="top"><font color="#000080" size="1"><font face="Arial">
		<%
		response.write Key & ": "
		%>
       </font></font>
		</td>
		<td><font color="#000080" size="1"><font face="Arial">
		<%		
		arrayname = Application.Contents(Key)
		for i = 0 to ubound(arrayname,2)
			response.write arrayname(0,i) & ", "
		next		
		%>
       </font></font>
		</td>
		</tr>
		<%
	Else
		If instr(ucase(Key),"PASSWORD") = 0 and instr(ucase(Key),"UID") then
		%>
		<tr>
		<td width="25%" valign="top"><font color="#000080" size="1">
        <font face="Arial">
		<%
		response.write Key & ": "
			%>
        </font></font>
			</td>
			<td><font color="#000080" size="1">
            <font face="Arial">
			<%
			Response.Write(Application.Contents(Key))
			%>
            &nbsp;
            </font></font>
			</td>
			</tr>
			<%
		End If	
	End If
Next
%>
</table>
  </center>
<%
End IF

If DBType = "ALL" OR dbtype = "S" then
%>
  <center>
<table border="2" width="90%" cellspacing="0" cellpadding="0">
	<tr>
	<td colspan="2" align="center" bgcolor="#99CCFF"><b><font color="#000080" size="1" face="Arial">SESSION CONTENTS VARIABLES</font></b></td>
	</tr>
<%
For Each Key in Session.Contents
	If Instr(ucase(Key),"RUNTIME") = 0 and instr(ucase(Key),"PASSWORD") = 0 then
		If ucase(Left(Key,3)) = "ARR" then    'We have an array 
			%>
			<tr>
			<td width="25%" valign="top"><font color="#000080" size="1"><font face="Arial">
			<%
			response.write Key & ": "
			%>
			</font></font></td>
			<td><font color="#000080" size="1"><font face="Arial">
			<%		
			arrayname = Session.Contents(Key)
			for i = 0 to ubound(arrayname,2)
				response.write arrayname(0,i) & ", "
			next		
			%>
          </font></font></td>
			</tr>
			<%
		Else
			%>
			<tr>
			<td width="25%" valign="top"><font color="#000080" size="1">
	        <font face="Arial">
			<%
			response.write Key & ": "
			%>
	       </font>
			</td>
			<td><font color="#000080" size="1">
          	 <font face="Arial">
			<%
			Response.Write(Session.Contents(Key))
			%>
          &nbsp;
          </font>
			</td>
			</tr>
			<%
		End If
	End IF 
Next
%>
</table>
  </center>

<%
End IF

If DBType = "ALL" OR dbtype = "F" then
%>

  <center>
<table border="2" width="90%" cellspacing="0" cellpadding="0">
	<tr>
	<td colspan="2" align="center" bgcolor="#99CCFF"><b><font color="#000080" size="1" face="Arial">FORM VARIABLES</font></b></td>
	</tr>
<%
For Each Key in Request.Form

%>
		<tr>
		<td width="25%"><font color="#000080" size="1">
        <font face="Arial">
		<%
		response.write Key & ": "
			%>
        </font>
			</td>
			<td><font color="#000080" size="1">
            <font face="Arial">
			<%
				Response.Write(Request.Form(Key))
			%>
            &nbsp;
            </font>
			</td>
			</tr>
			<%
Next
End IF

If DBType = "ALL" OR dbtype = "Q" then
%>

  <center>
<table border="2" width="90%" cellspacing="0" cellpadding="0">
	<tr>
	<td colspan="2" align="center" bgcolor="#99CCFF"><b><font color="#000080" size="1" face="Arial">QUERY STRING VARIABLES</font></b></td>
	</tr>
<%
For Each Key in Request.QueryString
%>
		<tr>
		<td width="25%"><font color="#000080" size="1">
        <font face="Arial">
		<%
		response.write Key & ": "
			%>
        </font>
			</td>
			<td><font color="#000080" size="1">
            <font face="Arial">
			<%
			Response.Write( Request.QueryString(Key))
			%>
            &nbsp;
            </font>
			</td>
			</tr>
			<%
Next
End IF

If DBType = "ALL" OR dbtype = "C" then
%>

  <center>
<table border="2" width="90%" cellspacing="0" cellpadding="0">
	<tr>
	<td colspan="2" align="center" bgcolor="#99CCFF"><b><font color="#000080" size="1" face="Arial">COOKIE VARIABLES</font></b></td>
	</tr>
<%
For Each Key in Request.Cookies
%>
		<tr>
		<td width="25%"><font color="#000080" size="1">
        <font face="Arial">
		<%
		response.write Key & ": "
			%>
        </font>
			</td>
			<td><font color="#000080" size="1">
            <font face="Arial">
			<%
			Response.Write( Request.Cookies(Key))
			%>
            &nbsp;
            </font>
			</td>
			</tr>
<%
Next
End IF

If DBType = "ALL" OR dbtype = "C" then
%>

  <center>
<table border="2" width="90%" cellspacing="0" cellpadding="0">
	<tr>
	<td colspan="2" align="center" bgcolor="#99CCFF"><b><font color="#000080" size="1" face="Arial">CLIENT CERTIFICATE VARIABLES</font></b></td>
	</tr>
<%
For Each Key in Request.ClientCertificate
%>
		<tr>
		<td width="25%"><font color="#000080" size="1">
        <font face="Arial">
		<%
		response.write Key & ": "
			%>
        </font>
			</td>
			<td><font color="#000080" size="1">
            <font face="Arial">
			<%
			Response.Write( Request.ClientCertificate(Key))
			%>
            &nbsp;
            </font>
			</td>
			</tr>
			<%
Next
End IF

If DBType = "ALL" OR dbtype = "S" then
%>

  <center>
<table border="2" width="90%" cellspacing="0" cellpadding="0">
	<tr>
	<td colspan="2" align="center" bgcolor="#99CCFF"><b><font color="#000080" size="1" face="Arial">SERVER VARIABLES</font></b></td>
	</tr>
<%
For Each Key in Request.ServerVariables
%>
		<tr>
		<td width="25%"><font color="#000080" size="1">
        <font face="Arial">
		<%
		response.write Key & ": "
			%>
        </font>
			</td>
			<td><font color="#000080" size="1">
            <font face="Arial">
			<%
				Response.Write(Request.ServerVariables(Key))
			%>
            &nbsp;
            </font>
			</td>
			</tr>
			<%
Next

End If
%>
</font></body>








