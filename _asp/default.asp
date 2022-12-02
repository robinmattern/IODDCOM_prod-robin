<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="incExpires.asp"-->
<!--#INCLUDE FILE="inccreateconnection.asp"-->

<%
Session("Visitor")=GetCounter()
if trim(session("sessionid"))="" then
	session("sessionid") = session.sessionid
end if	

'******************************************************
Function GetCounter()

SQLstr = "Select * from tConfiguration where Description = 'VisitorCount'"
Set rs = Conn.Execute(SQLStr)

' Increment Counter
LastCount = clng(rs("Settings"))
X = LastCount + 1

sqltext = "UPDATE tConfiguration SET Settings = '" & X & "' WHERE  Description = 'VisitorCount'"	
Conn.Execute(sqltext)
rs.Close
SET rs = Nothing

GetCounter = X

End Function
'******************************************************
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>Institute of Database Developers</title>
<!--#INCLUDE FILE="incheader.asp"-->
</head>

<body>
<!--#INCLUDE FILE="incbodyline.asp"-->
<div align="center" style="width: 705; height: 445">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="80%" style="border-collapse: collapse" bordercolor="#111111" align="left">
    <tr>
      <td valign="top" width="25%">
        <div align="center">
        <table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
          <tr>
            <td valign="top" align="center" height="25"><b><a href="RptMemberList.asp"><font face="Arial" color="#0000FF" size="2">Member
              Listing</font></a></b>
            </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><b><a href="RptMemberBio.asp"><font face="Arial" color="#0000FF" size="2">Bios</font></a></b>
            </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><b><a href="RptMemberContact.asp"><font face="Arial" color="#0000FF" size="2">Contact
              Sheet</font></a></b>
            </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><b><a href="RptMemberProjectsList.asp"><font face="Arial" color="#0000FF" size="2">Projects
              Listing</font></a></b>
            </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><b><a href="RptMemberProjects.asp"><font face="Arial" color="#0000FF" size="2">Projects
              Detail</font></a></b>
            </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><b><a href="meetings.asp"><font face="Arial" color="#0000FF" size="2">Next
              Meeting</font></a></b>
            </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><font face="Arial" size="2" color="#0000FF">&nbsp;<b><font face="Arial" color="#FFFFFF" size="3">
              </font>
              </b></font>
              <b>
              <a href="IODDMeeting.pdf"><font face="Arial" size="2" color="#0000FF">
              Meeting Location</font></a></b><font face="Arial" size="2" color="#0000FF">&nbsp;</font>
 </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><b><a href="othello.asp"><font face="Arial" color="#0000FF" size="2">Play
              Othello</font></a></b>
 </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><b>
            <a href="wonder.asp">
            <font face="Arial" color="#0000FF" size="2">Wonder </font></a> </b>
 </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><b>
            <a href="illusions.asp">
            <font face="Arial" color="#0000FF" size="2">Illusions </font></a> </b>
 </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25">&nbsp;
 </td>
          </tr>
          <tr>
            <td valign="top" align="center" height="25"><b><a href="SystemAction.asp?type=gurumail"><font face="Arial" color="#0000FF" size="2">Guru
              Mail</font></a></b>
 </td>
          </tr>
        </table>
        </div>
        </td>
      <center>
      <td>
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="90%" style="border-collapse: collapse" bordercolor="#111111">
            <tr>
              <td align="center">
          <ul>
            <li>
          <p align="left"><strong style="font-weight: 400"><font size="4">IODD 
          members have demonstrated
          expertise in one or more areas of database analysis, design,
          implementation, management and maintenance.  </font></strong></p>
              </li>
            <li>
          <p align="left"><strong style="font-weight: 400"><font size="4">IODD 
          members serve as advisors to other members and their companies about the constantly changing
          database issues.</font></strong></p>
              </li>
          </ul>
              </td>
            </tr>
            <tr>
              <td align="center"><strong>&nbsp;&nbsp;&nbsp;</strong></td>
            </tr>
            <tr>
              <td align="center">&nbsp;</td>
            </tr>
            <tr>
              <td align="center"><strong>&nbsp;&nbsp;</strong></td>
            </tr>
            <tr>
              <td align="center">
        <div align="center">
          <table border="0" bgcolor="#FFFFFF">
            <tr>
              <td bgcolor="#FFFFFF"><font size="2"><b>Contact</b><font color="#FFFF00" size="3">:</font></font></td>
              <td bgcolor="#FFFFFF"><a href="mailto:btroutma@iodd.com"><b>
              <font color="#0000FF" size="2">Bruce
                Troutman</font></b></a></td>
            </tr>
          </table>
        </div>
              </td>
            </tr>
          </table>
        </div>
        <hr width="70%">
        </td>
      </tr>
    <tr>
      <td valign="top" width="25%">
        &nbsp;</td>
      <td>
        &nbsp;</td>
      </tr>
    <tr>
      <td valign="top" width="25%">
        <p align="center"><strong><font size="2">You are visitor <%=Session("Visitor")%></font>&nbsp;</strong></td>
      <td>
        <p align="center"><font size="2">We use and recommend 
		<a href="http://www.icewarp.com"><font color="#0000FF">IceWarp Unified 
		Communications</font></a></font><font color="#0099CC"><br>
</font> </td>
      </tr>
    </table>
  </center>
  </center>
</div>

</body>

</html>