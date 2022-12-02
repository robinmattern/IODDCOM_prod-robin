<%@ LANGUAGE="VBScript" %> <% Response.Buffer = True %>
<!--#INCLUDE FILE="inccreateconnection.asp"-->

<% If len(trim(request("btn"))) > 0 then

		' "MARK COMPLETE"
		MarkComplete = Now()
		SQLStr = "UPDATE tdatacall Set MarkedCompleteat = '"& now() &"', MarkedCompleteBy = '"& replace(Session("ChangedBy"),"'","''") &"' WHERE datacallID = "& Session("ID") 
'		response.write SQLStr
'		response.end
		Conn.Execute(SQLStr)
		Session("valCount") = 0
		response.redirect "form_aap.asp?t=tdatacall" 
End If%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Validation</title>
</head>

<!--#INCLUDE FILE="incbodyline.asp"-->
<!--#INCLUDE FILE="incheader.asp"-->
<!--#INCLUDE FILE="incnav.asp"-->
<!-- #INCLUDE FILE = "incformlistaction.asp" -->
<body text="#000000" bgcolor="#F5F5F5">

<div align="center">
  <center>
  <p><br>
  </p>
  <div align="left">
  <table border="0" cellpadding="0" cellspacing="0" bordercolor="#000080">
    <tr>
      <td bgcolor="#CCCCFF">
      <p align="left"><font size="5"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Congratulations.&nbsp;</b></font><p align="left">
		<font size="5"><b>All validation rules have 
      passed.</b></font> </td>
    </tr>
    <tr>
      <td align="left">
      <p >&nbsp;<b><br>
      <br>
      <br>
      When you have completed the record and are ready to make it Read 
      Only, </b>
		<p></p><b>press the Mark Complete button below.</b>&nbsp;<br>
      <br>
      <br>
 </td>
    </tr>
    <tr> <td>
<form method="POST" action="_formvalidate.asp">
	<input type="submit" value="Mark Complete" name="BTN" style="float: left">
</form>
 </td>
    </tr>
  </table>
  </div>
  </center>
</div>
<p align="center"><br>&nbsp;
</p>

</body>
<!--#INCLUDE FILE="incfooter.asp"-->

</html>

</html>