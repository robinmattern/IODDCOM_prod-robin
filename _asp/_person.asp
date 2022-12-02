<%@ LANGUAGE="VBScript" %>
<!--#INCLUDE FILE="_incsessioncheck.asp"--> 
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>

<%'----- Expire the Page and Check that the User is currently Logged In. %>
<!--#INCLUDE FILE="_incExpires.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>MTS Users</title>
</head>

<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->
<!--#INCLUDE FILE="_incNAV.asp"-->
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<%
If len(trim(session("changedby"))) = 0 then
	response.redirect "default.asp"
End if 

'GET ARRAYS
	'Get Agencies
SQLStr = "SELECT lkupvalue FROM tlkup Where lkuptype='Organization' ORDER BY lkupvalue"
	'Response.write SQLStr
	'Response.end
	Set rs = Conn.Execute(SQLStr)
	arrOrganization = Null
	arrOrganization = rs.GetRows()
	Session("arrOrganization") = arrOrganization

PersonID = Request("PersonID")
Session("PersonID") = PersonID
SQLStr = "SELECT * FROM tperson WHERE PersonID = " & PersonID & ""
'response.write SQLStr
'response.end
Set rs = Conn.Execute(SQLStr)

%>
<form method="POST" action="_personaction.asp?from=_person">
 <center>
  <div align="left">
  <table cellspacing="0" cellpadding="0" border="0" bordercolor="#000080" style="border-collapse: collapse">
    <tr>
      <td valign="top" colspan="8">
        <p align="center"><font size="5"><b><%=session("project")%></b></font><b><font size="5"> Users</font></b></td>
    </tr>
</table>
	</div>
	<div align="left">
<table>
    <tr>
      <td valign="middle" colspan="9" height="50">
 		<input type="submit" value="Save Changes" name="btn" title="Save Changes">&nbsp;
       <input type="submit" value="Add A New User" name="btn" title="Add a New User">&nbsp;
       <input type="submit" value="Users List" name="btn" title="Go Back To User's List">      </tr>
</table></div>
	<font size ="2" color="red"><b><%=Session("valPersonAgencyMessage")%></b>

<div align="left">

<table border="1">
	<% RowCount = 0
	Do While Not rs.EOF
	   RowCount = Rowcount +1
	   %>

<tr>
    <input type="hidden" name="<%=RowCount%>PersonIDH" value="<%=rs("PersonID")%>">
	<td bgcolor="#000080" align="right" width="15%"><b><font color="#FFFFFF" size="2">Last Name</font></b></td>
	<td align="left"><font size ="2" color="red"><input type="text" name="<%=RowCount%>LastName" value="<%=rs("LastName")%>" size="18"></font></td>
</tr>
<tr>
	<td bgcolor="#000080" align="right" width="15%"><b><font color="#FFFFFF" size="2">First Name</font></b></td>
	<td align="left"><font size ="2" color="red"><input type="text" name="<%=RowCount%>FirstName"  value="<%=rs("FirstName")%>" size="18"></font></td>
</tr>
<tr>
	<td bgcolor="#000080" align="right" width="15%"><b><font color="#FFFFFF" size="2">Agency</font></b></td>
	<td align="left"><font size ="2" color="red">
		<select size="1" name="<%=RowCount%>Organization">
			<option value ="<%=Trim(rs("Organization"))%>" selected><%=Trim(rs("Organization"))%>
			<%for i = 0 to ubound(arrOrganization,2)
					if i = 0 AND Len(rs("Organization")) = 0 then%>
						<option value ="<% =arrOrganization(0,i) %>"><% =arrOrganization(0,i)%>
					<%else%>
		        		<%if arrOrganization(0,i)=rs("Organization") then%>
							<option value ="<% =arrOrganization(0,i) %>"><% =arrOrganization(0,i)%>
						<%else%>
							<option value ="<% =arrOrganization(0,i) %>"><% =arrOrganization(0,i)%>
						<%end if%>
		           <%end if%>
		    <% next %>
		</select>
	</td>
</tr>
<tr>
	<td bgcolor="#000080" align="right" width="15%"><b><font color="#FFFFFF" size="2">Role</font></b></td>
	<td align="left"><font size ="2" color="red">
		<select size="1" name="<%=RowCount%>Role">
         <%If rs("Role") = "ADM" Then%>
         	<option selected>ADM</option>
			<option>POC</option>
			<option>REP</option>
			<option>RO</option>
         <%ElseIf rs("Role") = "POC" Then%>
         	<option selected>POC</option>
			<option>ADM</option>
			<option>REP</option>
			<option>RO</option>
         <%ElseIf rs("Role") = "REP" Then%>
         	<option selected>REP</option>
			<option>ADM</option>
			<option>POC</option>
			<option>RO</option>
         <%ElseIf rs("Role") = "RO" Then%>
         	<option selected>RO</option>
			<option>ADM</option>
			<option>POC</option>
			<option>REP</option>
		  <%Else%>
         	<option selected>RO</option>
			<option>ADM</option>
			<option>POC</option>
			<option>REP</option>
		  <%End If%>
         </select>
	</td>
</tr>
<tr>
	<td bgcolor="#000080" align="right" width="15%"><b><font color="#FFFFFF" size="2">Email<br>
      &nbsp;(Used As Logon)</font></b></td>
	<td align="left"><font size ="2" color="red"><input type="text" name="<%=RowCount%>Logon" value="<%=rs("Logon")%>" size="33"></font></td>
</tr>
<tr>
	<td bgcolor="#000080" align="right" width="15%"><b><font color="#FFFFFF" size="2">Password</font></b></td>
	<td align="left"><font size ="2" color="red"><input type="text" name="<%=RowCount%>Password" value="<%=rs("Password")%>" size="15"></font></td>
</tr>
<tr>
	<td bgcolor="#000080" align="right" width="15%"><b><font color="#FFFFFF" size="2">Phone</font></b></td>
	<td align="left"><font size ="2" color="red"><input type="text" name="<%=RowCount%>Phone" value="<%=rs("Phone")%>" size="18"></font></td>
</tr>

<tr>
	<td bgcolor="#000080" align="right" width="15%"><b><font color="#FFFFFF" size="2">A</font></b></font><b><font color="#FFFFFF" size="2">ctive</font></b>	</td>
	<td align="left"><font size ="2" color="red">
		<select size="1" name="<%=RowCount%>ActiveUser">
     	  <%if rs("ActiveUser") <> "Yes" Then %>
		      <option>Yes</option>
		      <option selected>No</option>
		  <%else%>    
		      <option selected>Yes</option>
		      <option>No</option>
		  <%end If %>
       </select>
	</td>
</tr>
<tr>
	<td bgcolor="#000080" align="right" width="15%"><b><font color="#FFFFFF" size="2">Delete</font></b></td>
	<td align="left"><font size ="2" color="red"><input type="checkbox" name="<%=RowCount%>ckDelete" value="X" <%=checkedDelete%>></font></td>
</tr>

<% rs.MoveNext
Loop %>
<input type="hidden" name="TotalRowCount" value=<%=RowCount%>>
</table>
</div>
</form>
<!--#INCLUDE FILE="_incfooter.asp"-->
</body>
</html>