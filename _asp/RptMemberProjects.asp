<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>

<!--#INCLUDE FILE="inccreateconnection.asp"-->
<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">
<title>Institute of Database Developers - Member Projects</title>
<meta name="Microsoft Theme" content="none, default">
</head>

<%
' Get Info from Calling Page
First = trim(Request.QueryString("FirstName"))
Last = trim(Request.QueryString("LastName"))
ProjectCount = trim(Request.QueryString("ProjectCount"))
If ProjectCount = "" then 
	ProjectCount = "10"
End If	
If Last = "" then ' list everything
	SQLStr = "SELECT tProjectList.Dates, tProjectList.ProjectName, tProjectList.ProjectWeb, tProjectList.Client, tProjectList.ClientWeb, tProjectList.Location, tProjectList.Role, tProjectList.Duration, tProjectList.ProjectType, tProjectList.Industry, tmember.*, tProjectList.Description FROM tProjectList INNER JOIN tMember ON tProjectList.MemberID = tMember.MemberID WHERE Active = 'Y' ORDER BY tMember.LastName, tMember.FirstName, tProjectList.SortNumber DESC"
else
	SQLStr = "SELECT tProjectList.Dates, tProjectList.ProjectName, tProjectList.ProjectWeb, tProjectList.Client, tProjectList.ClientWeb, tProjectList.Location, tProjectList.Role, tProjectList.Duration, tProjectList.ProjectType, tProjectList.Industry, tmember.*, tProjectList.Description FROM tProjectList INNER JOIN tMember ON tProjectList.MemberID = tMember.MemberID WHERE tMember.LastName = '"& Last &"' and tMember.FirstName = '"&First&"' ORDER BY tMember.LastName, tMember.FirstName, tProjectList.sortnumber DESC"
end if
'response.write SQLStr
'response.end
set rs = conn.Execute(SQLStr)
%>
<%If Last = "" then %>
<!--#INCLUDE FILE="_incstrconvert.asp"-->
<!--#INCLUDE FILE="incheader.asp"-->
<%End If%>

<body>
<!--#INCLUDE FILE="incbodyline.asp"-->
<div align="left">
<table border="0" width=90%>
<%Member = "" %>
<% do 
	IF rs.EOF THEN
		Exit Do
	End If
	If Member <> rs("FirstName")+" "+rs("LastName") then %>	 
		<tr>
   		<td colspan = 3 align="center" valign="top">
        <p align="left">&nbsp;</p>
        <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" align="left">
          <tr>
            <td><font size="6"><%=rs("FirstName")+" "+rs("LastName")%></font>&nbsp;</td>
            <td rowspan="2" valign="top"><%=rs("Bio")%>&nbsp;</td>
          </tr>
          <tr>
            <td>
            <table align="center" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
              <tr>
                <td width="100%"><font size="2"><%=rs("Address1")%></font></td>
              </tr>
              <tr>
                <td width="100%"><font size="2"><%=rs("City")+", "+rs("State")+" "+rs("Zip")%></font></td>
              </tr>
              <tr>
                <td width="100%"><font size="2"><%=rs("Phone1")%></font></td>
              </tr>
              <tr>
                <td width="100%"><font size="2"><a href="mailto:<%=strconvert(rs("Email"),"html")%>">Email Address</a></font></td>
              </tr>
            </table>
		
		    </td>
          </tr>
        </table>
        <p align="left"><strong><font size="5">&nbsp;</font></strong></td>
		</tr>
		<tr>
   		<td> &nbsp; </td>
		</tr>
		<% Member = rs("FirstName")+" "+rs("LastName")
	End if%>

	<tr>
		<td width=45% valign="top">
			<table width=100%>
				<tr>
					<td valign="top"><small>Project: <%=rs("ProjectName")%></small></td>
				</tr>
				<%If rs("ProjectWeb") <> "" Then %>
				<tr>
					<td valign="top"><small>Project Web: <A HRef="<%=rs("ProjectWeb")%>"><%=rs("ProjectWeb")%></a></small></td>
			   </tr>
				<%End If%>
				<tr>
					<td valign="top"><small>Owner: <%=rs("Client")%></small></td>
			   </tr>
				<%If rs("ClientWeb") <> "" Then %>
				<tr>
					<td valign="top"><small>Owner Web: <A HRef="<%=rs("ClientWeb")%>"><%=rs("ClientWeb")%></a></small></td>
			   </tr>
				<%End If%>
				<%If rs("Dates") <> "" Then %>
				<tr>
					<td valign="top"><small>Dates: <%=rs("Dates")%></small></td>
			   </tr>
				<%End If%>
				<tr>
					<td valign="top"><small>Location: <%=rs("Location")%></small></td>
			   </tr>
				<tr>
					<td valign="top"><small>Role: <%=rs("Role")%></small></td>
			   </tr>
				<tr>
					<td valign="top"><small>Duration: <%=rs("Duration")%></small></td>
			   </tr>
				<tr>
					<td valign="top"><small>Type: <%=rs("ProjectType")%></small></td>
			   </tr>
			</table>
		</td>		   
    <td width=55% valign="top"><small><%=rs("Description")%></small></td>
  </tr>
  <tr>
	<td> &nbsp; </td>
  </tr>
 	<%rs.MoveNext%>
<%loop%>
	</table>
</div>

</body>
</html>