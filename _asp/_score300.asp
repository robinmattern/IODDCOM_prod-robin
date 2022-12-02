<% Response.Buffer = True %>

<html>
<%PageTitle = "Score 300"%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=PageTitle%></Title>
</head>
<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->
<!--#INCLUDE FILE="_incnav.asp"-->
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<!--#INCLUDE FILE="_increadconfiguration.asp"-->
<%
ReadConfiguration()
%>

<table border="2" width="90%" cellpadding="2" cellspacing="0" bordercolor="#000080" style="border-collapse: collapse">
  <tr>
    <td><p align="left"><em><strong><font color="#000080">
    <font size="5">Score 300 </font>(Total Updated + Total Completed&nbsp; = 300)</font></strong></em><font color="#FFFFFF"> 
    </font><font color="#000080"><i><font size="2">Not related to Exhibit 300</font></i></font><div align="center">
        <table border="0" cellpadding="2" cellspacing="0" height="10" align="left">
          <tr>
            <td bgcolor="#000080"><font size="1" color="#FFFFFF"><b>Scoring Legend&nbsp;</b></font></td>
            <td bgcolor="#FFC8CB"><font size="1"><b>&nbsp; 0&nbsp;&nbsp;</b></font></td>
            <td></td>
            <td bgcolor="#FFF09D"><font size="1"><b>1-299</b></font></td>
            <td></td>
            <td bgcolor="#90EE90"><font size="1"><b>&nbsp;300&nbsp;</b></font></td>
            <td></td>
          </tr>
        </table>
    </td>
  </tr>
</table>
&nbsp;


<%
SQLStr = "spscore300 '" & session("CurrentReportCYMO") & "'" 
'response.write sqlstr
'response.end
Set rs = Conn.Execute(SQLStr)
On Error Resume Next
fFirstPass = True
Do
    If rs.EOF Then Exit Do
    If Not fFirstPass Then
        rs.MoveNext
    Else
%>

   <table border="2" width="90%" cellspacing="0" bgcolor="#99CCFF" bordercolor="#000080" cellpadding="2" style="border-collapse: collapse">
  <tr>
    <td  align="center" height="30"><small>
    <font color="#FFFFFF"><b>Agency</b></font></small><font color="#FFFFFF">
    </font>
    </td>
    <td align="center" height="30" bgcolor="#000080"><small>
    <font color="#FFFFFF"><b>Score</b></font></small><font color="#FFFFFF">
    </font>
    </td>
    <td align="center" height="30" bgcolor="#000080"><small>
    <font color="#FFFFFF"><b>Last Update</b></font></small><font color="#FFFFFF">
    </font>
    </td>
    <td align="center" height="30" bgcolor="#000080"><small>
    <font color="#FFFFFF"><b>Contact</b></font></small><font color="#FFFFFF">
    </font>
    </td>
  </tr>
<%        fFirstPass = False
    End If
    If rs.EOF Then Exit Do
   	TotalScore = rs("TotalScore")
	Color = "LightGreen"
	if isnull(rs("LastUpdate")) then
	   	TotalScore = 0
		LastUpdated = "-"
		Color = "Pink"
	Else	
	   	TotalScore = 150
		LastUpdated = rs("LastUpdate")
		Color = "LightYellow"
	end if
%>
  <tr>
<% If TotalScore = 0 then %>
    <td bgcolor="<%=color%>"><font color="Black"><small><%=rs("Agency")%></small></font>&nbsp;</td>
<% Else %>
<% If TotalScore < 300 then %>
    <td bgcolor="<%=color%>"><font color="Black"><small><%	=rs("Agency")%></small></font>&nbsp;</td>
<% Else %>
<% If TotalScore = 300 then %>
    <td bgcolor="<%=color%>"><font color="Black"><small><%	=rs("Agency")%></small></font>&nbsp;</td>
<% Else %>
<% End If %>
<% End If %>
<% End If %>
    <td bgcolor="<%=color%>" align="left"><font color="Black"><small><%=TotalScore%></small></font>&nbsp;</td>
    <td bgcolor="<%=color%>" align="left"><font color="Black"><small><%=LastUpdated%></small></font>&nbsp;</td>
    <td bgcolor="<%=color%>" align="left"><font color="Black"><small><%=rs("Contact")%></small></font>&nbsp;</td>
  </tr>
<%Loop%>

</table>
   &nbsp;

</body>
<!-- End of the Body for this Page -->

<!-- Footer -->
<!--#INCLUDE FILE="_incFooter.asp"-->
</html>