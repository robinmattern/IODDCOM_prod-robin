<%@ LANGUAGE="VBScript" %>
<% Response.Buffer = True %>


<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->
<!--#INCLUDE FILE="_incnav.asp"-->
<!--#INCLUDE FILE="_inccreateconnection.asp"-->

<div align="left">
  <table border="1" cellpadding="2" cellspacing="0" bordercolor="#000080" bgcolor="#99CCFF" style="border-collapse: collapse">
    <tr>
      <td width="100%" align="center" bgcolor="#000080">

<p align="left"><b><font color="#FFFFFF" size="4">Help</font></b></p>

      </td>
    </tr>
    <tr>
      <td width="100%" align="left" bgcolor="#FFFFFF">

&nbsp;</td>
    </tr>

<%If Instr(session("Grouplist"),"ADM") > 0 OR Instr(session("Grouplist"),"POC") > 0 then %>
    <tr>
      <td width="100%" align="left" bgcolor="#FFFFFF">
<b><a href="InstructionsPOC.asp"><font size="3">Point of Contact Information</font></a></b></td>
    </tr>
<%End If%>
    <tr>
      <td width="100%" align="left" bgcolor="#FFFFFF">
<b><a href="InstructionsReporters.asp"><font size="3">Reporter Information</font></a></b></td>
    </tr>

<%If Instr(session("Grouplist"),"ADM") > 0 then %>
    <tr>
      <td width="100%" align="left" bgcolor="#FFFFFF">
<b><a href="nstructionsadministrators.asp"><font size="3">Administrator Information</font></a></b></td>
    </tr>
<%End If%>

    <tr>
      <td width="100%" align="left" bgcolor="#FFFFFF">
&nbsp;</td>
    </tr>
    
    <tr>
      <td width="100%" align="left" bgcolor="#FFFFFF">
&nbsp;</td>
    </tr>
    <tr>
<%If false then %>
      <td width="100%" align="left" bgcolor="#FFFFFF">
<b><a href="XMLExchangerInformation.asp"><font size="3">XML Exchanger Information</font></a></b></td>
    </tr>
    <tr>
      <td width="100%" align="left" bgcolor="#FFFFFF">
&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="left" bgcolor="#FFFFFF">
&nbsp;</td>
    </tr>
<%end if%>    
    </table>
</div>

<p align="center">

&nbsp;

<p align="center">

&nbsp;
</body>
<!--#INCLUDE FILE="_incfooter.asp"-->
</html>