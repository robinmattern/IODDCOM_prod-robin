<% 'Clear Variables
session("AgencyID") = 0
session("AgencyAbbr") = ""
session("AgencyName") = ""
If instr("ADM,POC",session("PersonRole")) = 0 then
	response.redirect "default.asp"
End If %>
<!--#INCLUDE FILE="_incsessioncheck.asp"-->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Admin Menu</title>
</head>

<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->
<!--#INCLUDE FILE="_incnav.asp"-->

<div align="center">
  <center>
  <table border="1" cellpadding="2" cellspacing="0" align="left" style="border-collapse: collapse">
    <tr>
      <td width="100%" align="center" bgcolor="#0000BB">

<b><font color="#FFFFFF" size="4">Administrative Menu</font></b>

      </td>
    </tr>
    <tr>
      <td width="100%" align="center" bgcolor="#FFFFFF">
&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center" bgcolor="#FFFFFF">
<b><a href="_personlist.asp"><font color="#000080" size="3">Add / Update Users</font></a></b>
      </td>
    </tr>
    <tr>
      <td width="100%" align="center" bgcolor="#FFFFFF">
&nbsp;</td>
    </tr>

<%If instr(session("grouplist"),"ADM") then %>
    <tr>
      <td width="100%" align="center" bgcolor="#FFFFFF">
<b><a href="_getemails.asp?level=STAFF"><font color="#000080" size="3">Get Email Addresses</font></a></b></td>
    </tr>
<%End If%>

<%If instr(session("grouplist"),"POC") then %>
    <tr>
      <td width="100%" align="center" bgcolor="#FFFFFF">
<b><a href="_getemails.asp?level=POC"><font color="#000080" size="3">Get Email 
Addresses</font></a></b></td>
    </tr>
<%End If%>

<%If instr(session("grouplist"),"ADM") then %>
    <tr>
      <td width="100%" align="center" bgcolor="#FFFFFF">

&nbsp;</td>
    </tr>
<%End If%>

  </table>
  </center>
</div>

<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>

<p>&nbsp;</p>

<p align="center">

&nbsp;

<div align="left">
  <table border="0" height="27" style="border-collapse: collapse" 
bordercolor="#111111" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center">
        &nbsp;&nbsp;
        <img border="0" src="_head_flagsm.jpg">
      </td>
    </tr>
    <tr>
      <TD valign="top" align="center">
      <div align="center">
        <table border="0" cellpadding="2" cellspacing="0">
          <tr>
            <td valign="top" rowspan="2">
            <p align="left"><font size="1">Visitor&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br></font>
            </td>
            <td valign="top">
            <p align="center"><font face="Arial" size="1">
&nbsp;This site is brought to you by <a href="http://www.gsa.gov">GSA</a> and <a href="http://www.8020data.com"> 80/20 Data Company</a><br>
      &nbsp;
Support: <a href="mailto:kennett.fussell@gsa.gov?subject=Support for Measurement Tracking System">Dr. Kennett Fussell</a>&nbsp;&nbsp;<a href="mailto:bruce.troutman@8020data.com?subject=Support For Measurement Tracking System">Bruce Troutman</a>&nbsp;&nbsp;<a href="mailto:evantage@aol.com?subject=Support for Measurement Tracking System">Richard Schinner
            </a>&nbsp;<a href="mailto:robin@inetpurchasing.com?subject=Support for Measurement Tracking System">Robin 
            Mattern</a></font></td>
            <td valign="top" rowspan="2"><font face="Arial" size="1">
            <a href="http://gsa.gov">
            <img border="0" src="_gsa_logo.gif" align="right" 
            alt="General Services Administration"></a></font></td>
          </tr>
          <tr>
            <td valign="top">
            <p align="center"><font face="Arial" size="1">
            <a target="_blank" href="_securityprivacy.asp">Security and Privacy Notice</a></font></td>
          </tr>
        </table>
      </div>
  </TD>

    </tr>
  </table>
</div>
</body>
</html>