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
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<%
'Fill OEProgram Array
SQLStr="SELECT RTrim([LkUpValue]) AS LkUp FROM tLkUp WHERE tLkUp.LkUpType='OEProgram' ORDER BY tLkUp.LkUpValue"
'Response.write SQLStr
Set rs = Conn.Execute(SQLStr)
arrOEProgram = Null
If rs.eof then
	arrOEProgram= Null
Else
	arrOEProgram= rs.GetRows()
	Session("arrOEProgram") = arrOEProgram
End If
rs.Close
SET rs = Nothing
%>

<div align="center">
<center>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.FirstName.value == "")
  {
    alert("Please enter a value for the \"First Name\" field.");
    theForm.FirstName.focus();
    return (false);
  }

  if (theForm.LastName.value == "")
  {
    alert("Please enter a value for the \"Last Name\" field.");
    theForm.LastName.focus();
    return (false);
  }

  if (theForm.Phone.value == "")
  {
    alert("Please enter a value for the \"Phone (Area Code + Phone Number + Extension)\" field.");
    theForm.Phone.focus();
    return (false);
  }

  if (theForm.Phone.value.length < 10)
  {
    alert("Please enter at least 10 characters in the \"Phone (Area Code + Phone Number + Extension)\" field.");
    theForm.Phone.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="_logon.asp?from=logonreg" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
<div align="left">
<table border="3" cellpadding="5" cellspacing="0">
    <tr>
   		<td valign="top" align="center" height="30">
          <div align="center">
            <table border="0" cellpadding="3">
<% 
'response.write "req-logon: " & session("logon")

				session("PersonPrefixName") = ""
				session("PersonFirstName") = ""
				session("PersonLastName") = ""
				session("PersonSuffixName") = ""
				session("PersonPhone") = ""
				session("OrgElement") = ""
				session("Program") = ""
%>
              <tr>
                <td align="right"><font size="2">Email</font></td>
                <td><b><%=session("logon")%></b></td>
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
                     <tr>
                        <td bgcolor="#FFFF99" width="489"><font size="2"><b>Please enter your 
						personal information</b></font></td>
                      </tr>
                    </table>
                  </div>
                  </td>
              </tr>
              <tr>
                <td align="left" valign="top" colspan="4">
                  <div align="center">
                    <center>
                    <table border="0" cellpadding="2" cellspacing="0">
                      <tr>
                        <td valign="top"><b>Prefix<br>
                          <input size="8" name="PrefixName" value="<%=session("PersonPrefixName")%>">
                          </b></td>
                        <td valign="top"><b>First<br>
                          <!--webbot bot="Validation" s-display-name="First Name" b-value-required="TRUE" --><input size="15" name="FirstName" value="<%=session("PersonFirstName")%>">
                  &nbsp;</b></td>
                        <td valign="top"><b>Last<br>
                          <!--webbot bot="Validation" s-display-name="Last Name" b-value-required="TRUE" --><input size="15" name="LastName" value="<%=session("PersonLastName")%>">&nbsp;</b></td>
                        <td valign="top"><b>Suffix<br>
                          <input size="8" name="SuffixName" value="<%=session("PersonSuffixName")%>">
                          &nbsp;</b></td>
                        <td valign="top"><b>Phone<br>
                          <!--webbot bot="Validation" s-display-name="Phone (Area Code + Phone Number + Extension)" b-value-required="TRUE" i-minimum-length="10" --><input size="15" name="Phone" value="<%=session("PersonPhone")%>"> 
                          &nbsp;</b></td>
                      </tr>
                    </table>
                    </center>
                  </div>
                </td>
              </tr>
              	<tr>
                <td align="left" valign="top" colspan="4">
                  <div align="center">
                    <center>
                    <table border="0" cellpadding="2" cellspacing="0" id="table1">
                      <tr>
                        <td valign="top"><b>Agency&nbsp;&nbsp;&nbsp;&nbsp; <br>
                          <%="DHS"%>
                          </b></td>
                        <td valign="top"><b>Organizational Element\Program<br>
                          &nbsp; </b>
	<select size="1" name="OEProgram">
	<%arrOEProgram = Session("arrOEProgram")
	   for i = 0 to ubound(arrOEProgram,2)%>
	      <option value="<% =arrOEProgram(0,i) %>"><% =arrOEProgram(0,i)%></option>
	 <% next %>
	 </select>
		       </font></td>
</td>
                        <td valign="top"><b>&nbsp; <br>
                          &nbsp;&nbsp; &nbsp;</b></td>
                        <td valign="top"><b>&nbsp;</b></td>
                        <td valign="top">&nbsp;</td>
                      </tr>
                    </table>
                    </center>
                  </div>
                </td>
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