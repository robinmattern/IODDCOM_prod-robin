<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>

<%'----- Expire the Page and Check that the User is currently Logged In. %>
<!--#INCLUDE FILE="incExpires.asp"-->



<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><!--#INCLUDE FILE="incTitle.asp"--> Message</title>
<meta name="Microsoft Border" content="none">
</head>
<!--#INCLUDE FILE="incbodyline.asp"-->
<body>
<div align="left">
  <table border="0" cellpadding="0" cellspacing="0" width="466">
    <tr>
      <td align="left" valign="middle">
        <p align="left"><font color="#3333FF" size="6">Message</font></p>
      </td>
    </tr>
    <tr>
      <td align="left" valign="middle">&nbsp;</td>
    </tr>
    <tr>
      <td align="left" valign="middle">
		<p align="left"><a href="../default.asp"><img border="0" src="../iodd.jpg"></a></td>
    </tr>
  <tr>
	<td valign="middle" align="left"><font size="5"><i><b>&nbsp;</b></i></font></td>
  </tr>
  <tr>
	<td valign="middle" align="left"><font size="5"><b><i><%=Session("Message")%></i></b></i></font></td>
  </tr>
  </table>
</div>
<table>
</table>
<p>&nbsp;</p>
</form>
</body>
</html>



