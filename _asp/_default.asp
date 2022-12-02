<%@ LANGUAGE="VBScript" %> <%Response.Buffer = True%> 
<html>

<!--#INCLUDE FILE="_increadconfiguration.asp"-->
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<%
systemabbr = "SAMP"
'Set Session("sessionid")
if trim(session("sessionid")) <> trim(session.sessionid)  or systemabbr <> session("SystemAbbr")then
	session("sessionid") = trim(session.sessionid)
	readconfiguration()  'Get variables for this application
end if	

'Down Message
'session("message")= "Sorry, but " & session("systemname") & " is unavailable for a short time.<BR>"
'response.redirect "_message.asp"
%>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=session("project")%>&nbsp; <%=session("systemname")%></title>
</head>

<body>

<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->
<!--#INCLUDE FILE="_incnav.asp"-->

<div align="left">
  <table border="0" cellpadding="0" cellspacing="0" bordercolor="#000080" style="border-collapse: collapse">
    <tr>
      <td>
      <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
          <tr>
            <td height="38" align="center">
            <table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1">
              <tr>
                <td align="center" height="30">
                <p align="left">&nbsp;</p>
                <div align="left">
                <table cellspacing="0" cellpadding="0" border="0">
                  <tr>
                    <td><b>
                    <font face="Times New Roman" size="2" color="#000080">This system 
                    allows all users to process&nbsp; 
                    via the Internet. </font></b></td>
                  </tr>
                  <b>
                  <tr>
                    <td><font face="Times New Roman" size="2">&nbsp;</font></td>
                  </tr>
                  <tr>
                    <td>
                    <ul>
                      <li>
                      <p align="left"><font face="Times New Roman" size="2">
                      Users can provide, verify and complete data via the Internet.</font>
                      </p>
                      </li>
                      <li>
                      <p align="left"><font face="Times New Roman" size="2">XML 
                      is used in this system.</font> </p>
                      </li>
                      <li>
                      <p align="left"><font face="Times New Roman" size="2">This 
                      system supports customers with IE, Netscape and Firefox browsers.</font>
                      </p>
                      </li>
                    </ul>
                    </td>
                  </tr>
                </table>
                </div>
                </b></td>
              </tr>
              </center>
              <tr>
                <td align="center" height="30" valign="top">&nbsp;</td>
              </tr>
              <tr>
                <td align="center" height="30" valign="top">
                <p align="left"><font size="2">Visitors, please contact the 
				Support Team for more information.</font></p>
                <p align="left"><font size="2">Reporters and 
				managers, 
                please logon.</font></p>
                </td>
              </tr>
              <tr>
                <td align="center">
                <p align="left">&nbsp;</p>
                </td>
                </font>
              </tr>
            </table>
            </td>
          </tr>
        </table>
        </center>
      </div>
      </td>
    </tr>
  </table>
</div>
<font face="CG Times">&nbsp; </font></font>

</body>
<!--#INCLUDE FILE="_incfooter.asp"-->
</html>