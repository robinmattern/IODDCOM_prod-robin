<%@ LANGUAGE="vbscript" %>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="_incsessioncheck.asp"-->
<!--#INCLUDE FILE="_incExpires.asp"-->

<%
If len(request("level")) = 0 then
	session("Message") = "An error has occured with getemails.asp. Please contact the support team."
	response.redirect "_message.asp"
End If
%>

<%PageTitle="Get Email Addresses"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<META NAME="Author" CONTENT="80/20 Data Co.">
<META NAME="Generator" CONTENT="Microsoft FrontPage 6.0">
<TITLE><%=PageTitle%></TITLE>
<TITLE>Annual Comprehensive Reports</TITLE>
</HEAD>
<!--#INCLUDE FILE="_incheader.asp"-->
<!--#INCLUDE FILE="_incnav.asp"-->
<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->

<br>
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<%If request("TYPE") > "" then 
If ucase(request("level")) = "STAFF" then
	Select Case request("Type")
	Case "Staff"
		SQLStr = " Select logon from tperson where role = 'ADM' and Activeuser <> 'No'"
	Case "POC"
		SQLStr = " Select logon from tperson where role = 'POC' and Activeuser <> 'No'"
	Case "POCNOLOG"
		SQLStr = " Select logon from tperson where role = 'POC' AND lastlogondate IS NULL and Activeuser <> 'No'"
	Case "REP"
		SQLStr = " Select logon from tperson where role =  'REP' and Activeuser <> 'No'"
	Case "REPNOLOG"
		SQLStr = " Select logon from tperson where role =  'REP' AND lastlogondate IS NULL and Activeuser <> 'No'"
	Case "POCREP"
		SQLStr = " Select logon from tperson where (role = 'POC' OR role = 'REP') and Activeuser <> 'No'"
	Case Else
		SQLStr = " Select logon from tperson where role <> 'RO' AND Activeuser <> 'No' ORDER by Logon"
	End Select
End If
If request("level") = "POC" then
	Select Case request("Type")
	Case "Staff"
		SQLStr = " Select logon from tperson where role = 'ADM' and Activeuser <> 'No'"
	Case "POC"
		SQLStr = " Select logon from tperson where role = 'POC' and Activeuser <> 'No' and Organization = '" & Session("AgencyAbbr") & "'"
	Case "REP"
		SQLStr = " Select logon from tperson where role =  'REP' and Activeuser <> 'No' and Organization = '" & Session("AgencyAbbr") & "'"
	Case "REPNOLOG"
		SQLStr = " Select logon from tperson where role =  'REP' and Activeuser <> 'No' AND lastlogondate IS NULL and Organization = '" & Session("AgencyAbbr") & "'"
	Case "POCREP"
		SQLStr = " Select logon from tperson where (role = 'POC' OR role = 'REP') and Activeuser <> 'No' and Organization = '" & Session("AgencyAbbr") & "'"
	Case Else
		SQLStr = " Select logon from tperson where Organization = '" & Session("AgencyAbbr")  & "' and Activeuser <> 'No'"
	End Select
End If
'response.write "ss: " & sqlstr
'response.end
Set rs = Conn.Execute(SqlStr) 
%>
<table border="1" width="90%" bordercolor="#000080" cellpadding="3" cellspacing="0">
<tr>
<td align=Center bgcolor="#000080"><b><font color="#FFFFFF" size="4"><%=request("Type")%> Email Addresses</font></b></td>
</tr>
<tr>
<td>
<%do while not rs.eof 
	response.write rs("logon") & "; "
	rs.movenext
loop%> &nbsp;</td>
</tr>
</table>		
<%End If %>
<p align="left"><font size="3">Get Email Addresses for:</font></p>

<table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left">
          <table border="0" cellpadding="0" cellspacing="0" height="135">

<%If request("level") = "STAFF" then %>
<tr>
	<td align="left" height="40">
    <p align="left">
    <a href="_getemails.asp?level=<%=request("level")%>&type=POC">
    POC's Only</a></p>
	</td>
</tr>
<tr>
	<td align="left" height="40">
    <p align="left">
    <a href="_getemails.asp?level=<%=request("level")%>&type=POCNOLOG">
    POC's Who have not logged on to the system</a></p>
	</td>
</tr>
<%End If%>

<tr>
	<td align="left" height="40">
    <p align="left">
    <a href="_getemails.asp?level=<%=request("level")%>&type=REP">
    REP's Only</a></p>
    </td>
</tr>
<tr>
	<td align="left" height="40">
    <p align="left">
    <a href="_getemails.asp?level=<%=request("level")%>&type=REPNOLOG">
    REP's Who have not logged on to the system</a></p>
	</td>
</tr>
<tr>
	<td align="left" height="40">
    <p align="left">
    <a href="_getemails.asp?level=<%=request("level")%>&type=POCREP">
    POCs and REP's</a>
    </p>
	</td>
</tr>

<tr>
	<td align="left" height="40">
    <p align="left"><a href="_getemails.asp?level=<%=request("level")%>&type=Staff">Staff</a>
    </p>
	</td>
</tr>

<%If request("level") = "STAFF" then %>
<tr>
	<td align="left" height="40">
    <p align="left">
    <a href="_getemails.asp?level=<%=request("level")%>&type=ALL">All Addresses</a>
    </p>
	</td>
</tr>
<%End If%>

          </table>

    </td>
  </tr>
</table>
<p align="left">
&nbsp;    

<!-- End of the Body for this Page -->
</BODY>
<!-- Footer -->
<!--#INCLUDE FILE="_incFooter.asp"-->
</html>