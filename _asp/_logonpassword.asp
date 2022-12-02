<%@ LANGUAGE="vbscript" %>
<% Response.Buffer = True %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<META NAME="Author" CONTENT="80/20 Data Co.">
<META NAME="Generator" CONTENT="Microsoft FrontPage 6.0">
<TITLE>Logon Info</TITLE>
</HEAD>
<body>
<!--#INCLUDE FILE="_incBodyLine.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->

<div align="center">
<center>
<form method="POST" action="_logon.asp?from=logonpassword" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
<div align="left">
<table border="3" cellpadding="5" cellspacing="0">
    <tr>
   		<td valign="top" align="center" height="30">
          <div align="center">
            <table border="0" cellpadding="3">
              <tr>
                <td align="right"><font size="2">Existing Logon</font></td>
                <td><b><%=Session("PersonLogon")%></b></td>
                <td><font size="2">Existing Password</font></td>
                <td><b><%=Session("Password")%></b></td>
              </tr>
            </table>
          </div>
	   </td>
    </tr>
  </center>
	<tr>
   		<td valign="top" align="center">
          <div align="center">
            <table border="0" cellspacing="1" cellpadding="0">
				<%If Len(Session("valPasswordError")) > 0 Then%>
              <tr>
                <td align="center" colspan="4"><font size"3" color="red"><b><%=Session("valPasswordError")%></b></font></td>
              </tr>
				<%End If%>
              <tr>
                <td align="center" colspan="4">
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="493">
                    </table>
                  </div>
                  </td>
              </tr>
              <tr>
                <td align="left" valign="top" colspan="4">
                  <div align="left">
                  </div>
                </td>
              </tr>
                <tr>
                <td align="right">
                  <p align="center"><b><font size="1">&nbsp;</font></b></td>
                <td colspan="3" align="right">
                  <p align="center"><font size="4">It is time to change your 
                  password.</font></td>
                </tr>
              <tr>
                <td align="right">
                  &nbsp;</td>
                <td colspan="3" align="right">
                  &nbsp;</td>
              </tr>
              <tr>
                <td align="right"><b>&nbsp;</b></td>
                <td align="right"><b>Enter new Password</b> </td>
                <td align="left"><input size="12" name="NewPassword"></td>
                <td align="left"> <font color="#FF0000" size="2"><b><%=Session("valPersonPassword")%><b></font></td>
              </tr>
              <tr>
                <td align="right"></td>
                <td align="right"><b>Confirm new Password&nbsp; </b> </td>
                <td align="left"><input size="12" name="ConfirmPassword"> </td>
                <td align="left"> </td>
              </tr>
              <tr>
                <td align="right"><font size="1">&nbsp; </font></td>
                <td colspan="3"><font size="1">&nbsp; </font></td>
              </tr>
            </table>
          </div>
	   </td>
    </tr>
    <tr>
      <td valign="top" align="center">
          <div align="center">
            <table border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td align="center" valign="top">
          <input type="submit" value="Save" name="btn"></td>
              </tr>
            </table>
          </div>
      </td>
    </tr>
  </table>
</div>
</div>
</form>
    <!-- End of the Body for this Page -->
</body>
<!-- Footer -->
<!--#INCLUDE FILE="_incFooter.asp"-->