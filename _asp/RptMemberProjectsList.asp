<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="_incstrconvert.asp"-->
<!--#INCLUDE FILE="inccreateconnection.asp"-->

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">
<title>Member Projects</title>
<meta name="Microsoft Theme" content="none, default">
</head>

<body>
<!--#INCLUDE FILE="incbodyline.asp"-->

<%
' Get Info from Calling Page
Title=trim(Request.QueryString("Title"))
First = trim(Request.QueryString("FirstName"))
Last = trim(Request.QueryString("LastName"))
Yr = Request("Year")
Yearclause = ""
If not isempty(Yr) then
	Yearclause = " AND Dates >= '" & Yr & "'"
End If	

'response.write last & "<BR>"
'response.write first & "<BR>"
'response.end

If "" & Last = "" then 
	SQLStr = "SELECT tProjectList.Dates as Projectdates, tProjectList.ProjectName, tProjectList.ProjectWeb, tProjectList.Client, tProjectList.ClientWeb, tProjectList.Location, tProjectList.Role, tProjectList.Duration, tProjectList.ProjectType, tProjectList.Industry, tProjectList.Description, tmember.*  FROM tProjectList INNER JOIN tMember ON tProjectList.MemberID = tMember.MemberID WHERE Active = 'Y' " & YearClause & " ORDER BY tMember.LastName, tMember.FirstName, tProjectList.SortNumber DESC"
else
	SQLStr = "SELECT tProjectList.Dates as Projectdates, tProjectList.ProjectName, tProjectList.ProjectWeb, tProjectList.Client, tProjectList.ClientWeb, tProjectList.Location, tProjectList.Role, tProjectList.Duration, tProjectList.ProjectType, tProjectList.Industry, tProjectList.Description, tmember.*  FROM tProjectList INNER JOIN tMember ON tProjectList.MemberID = tMember.MemberID WHERE tMember.LastName = '"& Last &"' and tMember.FirstName = '"& First &"'" & Yearclause & " ORDER BY Dates DESC"
end if
'response.write SQLStr &"<BR>"
'response.end
set rs = conn.Execute(SQLStr)

%>
<%If isempty(Title) then%>
<!--#INCLUDE FILE="incheader.asp"-->
<%Else%>
	<p><b><font size="4"><%=Title%></font></b></p>
<%End If%>


<table align = "left" border="1" width="80%" cellspacing="0" cellpadding="0" style="border-collapse: collapse">
<%MemberId = 0%>
<% do 
	IF rs.EOF THEN
		Exit Do
	End If
	If MemberId <> rs("MemberId") then 
	   MemberId = rs("MemberId") 
	   Name = rs("FirstName")+" "+rs("LastName")
	   Addr = rs("Address1")
	   CSZ = rs("City")+", "+rs("State")+" "+rs("Zip")
	   Phone = rs("Phone1")
	   Email = strconvert(rs("Email"),"html")
	   %>
		<tr>
   				<td colspan="6" bgcolor="#0066CC" >&nbsp;</td>
		</tr>
		<tr>
   				<td colspan="6" >
                  <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" align="left">
                    <tr>
                      <td><font size="6"><%=Name%>
                        </font>&nbsp;</td>
                      <td rowspan="2" valign="top"><%=rs("Bio")%>
                        &nbsp;</td>
                    </tr>
                    <tr>
                      <td>
                        <table align="center" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
                          <tr>
                            <td width="100%"><font size="2"><%=Addr%>
                              </font></td>
                          </tr>
                          <tr>
                            <td width="100%"><font size="2"><%=CSZ%>
                              </font></td>
                          </tr>
                          <tr>
                            <td width="100%"><font size="2"><%=Phone%>
                              </font></td>
                          </tr>
                          <tr>
                            <td width="100%"><font size="2"><a href="mailto:<%=Email%>">Email Address</a>
                              </font></td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                  <p>&nbsp;</td>
		</tr>
		<tr>
   				<td colspan="6" >
                  &nbsp;</td>
		</tr>
		<tr>
			<td valign="top" width="25%"><b><font face="Times New Roman" size="1">Project</font></b></td>
			<td valign="top" width="20%"><b><font face="Times New Roman" size="1">Role</font></b></td>
			<td valign="top" width="20%"><b><font face="Times New Roman" size="1">Type</font></b></td>
			<td valign="top" width="5%"><b><font face="Times New Roman" size="1">Dates</font></b></td>
			<td valign="top" width="5%"><b><font face="Times New Roman" size="1">Project Web</font></b></td>
			<td valign="top" width="15%"><b><font face="Times New Roman" size="1">Industry</font></b></td>
		</tr>
	<%End if%>
	   <tr>
				<td valign="top" ><font face="Times New Roman" size="1" ><%=rs("ProjectName")%></font>&nbsp;</td>
				<td valign="top" ><font face="Times New Roman" size="1"><%=rs("Role")%></font>&nbsp;</td>
				<td valign="top" ><font face="Times New Roman" size="1" ><%=rs("ProjectType")%></font>&nbsp;</td>
				<td valign="top" ><font face="Times New Roman" size="1" ><%=trim(rs("ProjectDates"))%></font>&nbsp;</td>
				<%If rs("ProjectWeb") <> "" Then %>
					<td valign="top" ><A HRef="<%=rs("ProjectWeb")%>"><font face="Times New Roman" size="1"><%=mid(rs("ProjectWeb"),8)%></font></a>&nbsp;</td>
				<%Else%>
	    			<td valign="top" ><font face="Times New Roman" size="1">&nbsp;</font></td>
				<%End If%>
				<td valign="top" ><font face="Times New Roman" size="1"><%=trim(rs("Industry"))%></font>&nbsp;</td>
      </tr>

 	<%rs.MoveNext%>
<%loop%>

</table>
</body>
</html>