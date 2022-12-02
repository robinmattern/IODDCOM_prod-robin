<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>
<%'----- Expire the Page and Check that the User is currently Logged In. %>
<!--#INCLUDE FILE="_incExpires.asp"-->
<!--#INCLUDE FILE="_incsessioncheck.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>EMail</title>
</head>

<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.To.value == "")
  {
    alert("Please enter a value for the \"valid email addreses \" field.");
    theForm.To.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ0123456789-@;._";
  var checkStr = theForm.To.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter, digit and \"@;._\" characters in the \"valid email addreses \" field.");
    theForm.To.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="_generateemail.asp" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
    <center>
    <div align="left">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td align="center" valign="top" colspan="3">
          <hr size="1" width="90%">
        </td>
      </tr>
      <tr>
        <td align="right" valign="top"><b>Email From:</b></td>
        <td align="right" valign="top">&nbsp;&nbsp;</td>
        <td >
  		<%If len(trim(session("EmailFrom"))) > 0  then %>
	       	<%=session("EmailFrom")%>
		<%Else%>
	  		<%If len(trim(session("PersonEmail"))) > 0 then %>
	       		<%=session("PersonEmail")%>
		       	<%session("EmailFrom")=session("PersonEmail")%>
			<%Else%>
	       		<%="cfomts.support@govwebs.net"%>
		       	<%session("EmailFrom")="cfomtssupport@fido.gov"%>
	       	<%End if%>
       	<%End if%>
		</td>
      </tr>
      <tr>
        <td colspan="2">&nbsp;</tr>
      <tr>
        <td align="right" valign="top" width="125"><b>Email To:<br>
          </b><font size="2" color="#0000FF">
          Enter as many Email<br>
          addresses as you want.<br>
          Separate addresses<br>
          with semicolons (;)<br>
        	No spaces</font>
          
        </td>
        <td align="right" valign="top"></td>
        <td valign="top" >
        <!--webbot bot="Validation" s-display-name="valid email addreses " s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" s-allow-other-chars="@;._" b-value-required="TRUE" --><textarea rows="5" name="To" cols="60"></textarea>
		</td>
      </tr>
      <tr>
        <td align="right" valign="top" colspan="2"><b>&nbsp;</b></td>
        <td >
  		&nbsp;</td>
      </tr>
      <tr>
        <td align="right" valign="top"><b>Subject:</b></td>
        <td align="right" valign="top"></td>
        <td >
  		<%If len(trim(session("EmailSubject"))) > 0 then %>
	       	<p><%=session("EmailSubject")%></p>
		<%Else%>
            <input type="text" name="Subject" size="70">
        <%End If%>
        </td>
      </tr>
      <tr>
        <td colspan="3" align="center">
          &nbsp;</td>
      </tr>
      <tr>
        <td align="right" valign="top"><b>Message:</b></td>
        <td align="right" valign="top"></td>
        <td >
  		<%If len(trim(session("EmailMessage"))) > 0 then %>
	       	<p><%=session("EmailMessage")%></p>
		<%Else%>
            <p><textarea rows="6" name="Message" cols="60"></textarea></p>
        <%End If%>
        </td>
      </tr>
    </center>
      <tr>
        <td colspan="3" align="center">
          <hr size="1" width="90%">
        </td>
      </tr>
      <tr>
        <td colspan="3" align="center"><input type="submit" value="Send Email" name="btn"></td>
      </tr>
    </table>
    </div>
</form>
</body>
</html>