<%@ LANGUAGE="VBScript" %>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="_incexpires.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Message</title>
</head>
<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->
<!--#INCLUDE FILE="_incnav.asp"-->

<div align="left">
  <table border="0" cellpadding="0" cellspacing="0" width="466">
    <tr>
      <td align="center" valign="middle">&nbsp;</td>
    </tr>
  <tr>
	<td valign="middle"><font size="4"><i><b>&nbsp;</b></i></font></td>
  </tr>
  <tr>
	<td valign="middle"><font size="4"><i><b>&nbsp;&nbsp;<%=Session("Message")%></i></b></i></font></td>
  </tr>
  </table>
</div>
<table>
</table>
<p>&nbsp;</p>

</body>
<!--#INCLUDE FILE="_incfooter.asp"-->
</html>

