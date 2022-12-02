<%@ LANGUAGE="VBScript" %>
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>
<%
response.write 'I AM STOPPING HERE!"
response.end
%>
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
'	SQLStr = "SELECT TOP "& ProjectCount & " tmember.*, tProjectList.Dates, tProjectList.ProjectName, tProjectList.ProjectWeb, tProjectList.Client, tProjectList.ClientWeb, tProjectList.Location, tProjectList.Role, tProjectList.Duration, tProjectList.ProjectType, tProjectList.Industry, tProjectList.Description FROM tProjectList INNER JOIN tMember ON tProjectList.MemberID = tMember.MemberID WHERE Active = 'Y' ORDER BY tMember.LastName, tMember.FirstName, tProjectList.SortNumber DESC"
else
'	SQLStr = "SELECT TOP "& ProjectCount & " tmember.*, tProjectList.Dates, tProjectList.ProjectName, tProjectList.ProjectWeb, tProjectList.Client, tProjectList.ClientWeb, tProjectList.Location, tProjectList.Role, tProjectList.Duration, tProjectList.ProjectType, tProjectList.Industry, tProjectList.Description FROM tProjectList INNER JOIN tMember ON tProjectList.MemberID = tMember.MemberID WHERE tMember.LastName = '"& Last &"' and tMember.FirstName = '"&First&"' ORDER BY tMember.LastName, tMember.FirstName, tProjectList.sortnumber DESC"
end if
response.write SQLStr
response.end
set rs = conn.Execute(SQLStr)
%>
<%If Last = "" then %>
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
	If Member <> rs("FirstName")+' '+rs("LastName") then %>	 
		<tr>
   		<td colspan = 3 align="center" valign="top"><strong><font size="5">Projects by <%=rs("FirstName")+' '+rs("LastName")%></font></strong></td>
		</tr>
		<tr>
   		<td> &nbsp; </td>
		</tr>
		<% Member = rs("FirstName")+' '+rs("LastName")
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
				<tr>
					<td valign="top"><small>Dates: <%=rs("Dates")%></small></td>
			   </tr>
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