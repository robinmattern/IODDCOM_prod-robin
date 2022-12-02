<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>


<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>EMail</title>
</head>

<!--#INCLUDE FILE="incbodyline.asp"-->
<!--#INCLUDE FILE="incheader.asp"-->
<!--#INCLUDE FILE="inccreateconnection.asp"-->
  <center>
  <div align="left">
  <table border="0" cellpadding="0" cellspacing="0" width="80%">
    <tr>
      <td>
        <p align="left"><b><font size="5" color="#0000FF">Guru Mail</font></b></td>
    </tr>
    <tr>
      <td>
        <p align="left"><b><font size="3" color="#0000FF">Send IODD members
        your comments and questions</font></b></td>
    </tr>
    <tr>
      <td>
        <p align="left">
        <font size="3" color="#0099CC">&nbsp;</font></td>
    </tr>
  </table>
	</div>
	<form method="POST" action="SystemAction.asp?from=email">
    <div align="left">
    <table border="0" cellpadding="0" cellspacing="0" width="620">
      <tr>
        <td align="right" width="71" valign="top"><b><font color="#0000FF">From:&nbsp;&nbsp;</font></b></td>
        <td width="393">
            <input type="text" name="EmailFrom" size="55">
		</td>
        <td width="150">
            <b>
            <font color="#0000FF"><font size="2">(Email Address)</font> </font>
		    </b>
		</td>
      </tr>
      <tr>
        <td align="right" width="71" valign="top"><b><font color="#0000FF">Subject:&nbsp;&nbsp;</font></b></td>
        <td width="545" colspan="2">
            <input type="text" name="Subject" size="55">
        </td>
      </tr>
      <tr>
        <td align="right" width="71" valign="top"><b><font color="#0000FF">Message:&nbsp; </font> </b></td>
        <td width="545" colspan="2">
            <p><textarea rows="14" name="Message" cols="60"></textarea></p>
        </td>
      </tr>
      <tr>
        <td colspan="3" align="center" width="618">
          <p align="center"><font color="#0099CC">
			<input type="submit" value="Send Email" name="btn" style="float: left"></font></td>
      </tr>
    </table>
  	</div>
  <p align="center"></p>
</form>


</html>