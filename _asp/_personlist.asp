<%@ LANGUAGE="VBScript" %>
<!--#INCLUDE FILE="_incsessioncheck.asp"-->
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% Response.Buffer = True %>
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Users</title>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
</head>
<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->
<!--#INCLUDE FILE="_incnav.asp"-->
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<!-- #INCLUDE FILE = "_incexpires.asp" -->
<!-- #INCLUDE FILE = "_incformlistaction.asp" -->


<%
If len(trim(session("changedby"))) = 0 then
	response.redirect "default.asp"
End if 

'Delete the Users records that still have "-Email Address" in the logon field
SQLStr = "DELETE tPerson FROM tPerson WHERE Logon = '-Email Address'"
Conn.Execute(SQLStr)
'response.write SQLStr
'response.end

if session("LastPage") <> "_personlist.asp" then
	Session("SortType") = "lastname + ' ' + firstname"
	Session("SortDirection") = "ASC"
end if	
Session("LastPage") = "_personlist.asp"
SqlStr = "SELECT Count(*) AS CountOfPersons FROM tperson"
'response.write sqlstr
'response.end
Set rs = Conn.Execute(SQLstr)
CountOfPersons = rs("CountOfPersons") 
PersonsBanner = CountOfPersons & " Users"

If session("DataEntryRole") = "POC" then
	PersonsBanner = session("Agencyabbr") & " - " & session("AgencyName")
end if

' Sorting -----
If len(trim(request("sort"))) <> 0 then

 select case lcase(request("sort"))
 case "name"
	SortType = "lastname + ' ' + firstname"
 case "organization"
	Sorttype = "Organization"
 case "active"
	Sorttype = "ActiveUser"
 case "role"
	Sorttype = "Role"
 case "logon"
	Sorttype = "Logon"
 case "phone"
	Sorttype = "Phone"
 case "lastlogon"
	Sorttype = "LastLogonDate"
 case else
	SortType = "lastname + ' ' + firstname"
	SortDirection = "ASC"	
 End select

 If session("SortType") = request("Sort") then
	if Session("SortDirection") = "ASC" then
		SortDirection = "DESC"
	else 
		SortDirection = "ASC"	
	end if	
 End If

Else

 If len(trim(session("SortType"))) = 0 then
	SortType = "lastname + ' ' + firstname"
	SortDirection = "ASC"	
 Else
    SortDirection = Session("SortDirection")
    SortType = Session("SortType")
 End If

End if

Session("SortDirection") = SortDirection
Session("SortType") = SortType

If session("DataEntryRole") = "POC" then
	SqlStr = "SELECT tPerson.* FROM tPerson WHERE Organization Like '" & session("PersonOrganization") & "%' ORDER BY " & session("Sorttype") & " " & session("sortdirection")
Else
	SqlStr = "SELECT * FROM tPerson ORDER BY " & session("Sorttype") & " " & session("sortdirection")
End if	
'response.write sqlstr
'response.end
Set rs = Conn.Execute(SQLstr)
%>

<form method="POST" action="_personaction.asp?from=_personlist" name="FrontPage_Form1" >

          <table>
			<tr>
				<td align="center">
                  <p align="center"><b><font size="4" color="#000080"><%=session("project")%> 
					Users&nbsp;&nbsp;&nbsp;&nbsp;<font size="2" color="#009933"><b><%=PersonsBanner%></font></font></b></p>
                </td>
		<TABLE border="1" bordercolordark="#C0C0C0" bordercolorlight="#DCDCDC" bordercolor="#C0C0C0" cellspacing="0">
			  <TR vAlign=top>
   		 		<TD bgColor="#000080" align="left" valign="middle" colspan="8" height="28">
      			<input type="submit" value="Add A New User" name="btn">
     			</TD>
  			  </TR>

            <tr>
              <td BGCOLOR="#000080" HEIGHT="16" align="center" valign="top"><b><a href="_personlist.asp?Sort=Name"><font face="Arial" color="#FFFFFF" size="2">
				Name</font></a></b></td>
            	<td BGCOLOR="#000080" HEIGHT="16" align="center" valign="top"><b><a href="_personlist.asp?sort=organization"><font face="Arial" color="#FFFFFF" size="2">
				Organization</font></a></b></td>
            	<td BGCOLOR="#000080" HEIGHT="16" align="center" valign="top"><b><a href="_personlist.asp?sort=active"><font face="Arial" color="#FFFFFF" size="2">
				Active</font></a></b></td>
            	<td BGCOLOR="#000080" HEIGHT="16" align="center" valign="top"><b><a href="_personlist.asp?sort=role"><font face="Arial" color="#FFFFFF" size="2">
				Role</font></a></b></td>
            	<td BGCOLOR="#000080" HEIGHT="16" align="center" valign="top"><b><a href="_personlist.asp?sort=logon"><font face="Arial" color="#FFFFFF" size="2">
				Logon (eMail)</font></a></b></td>
            	<td BGCOLOR="#000080" HEIGHT="16" align="center" valign="top"><b><a href="_personlist.asp?sort=indicators"><font face="Arial" color="#FFFFFF" size="2">
				Phone</font></a></b></td>
            	<td BGCOLOR="#000080" HEIGHT="16" align="center" valign="top"><b><a href="_personlist.asp?sort=lastlogon"><font face="Arial" color="#FFFFFF" size="2">
				Last Logon</font></a></b></td>
      		  </tr>
	<%
		do while NOT rs.EOF
		FullName = rs("LastName") & ", " & rs("FirstName")
		If instr(rs("Logon"),"@") > 0 Then
			Logon = "<a href = mailto:" & rs("Logon") & ">" & rs("Logon") & "</a>"
		Else
			Logon = rs("Logon")
		End If
		If Len(rs("Phone")) = 0 Or IsNull(rs("Phone")) Then
			Phone = "----------------"
		Else
			Phone = rs("Phone")
		End If
		If Len(rs("LastLogonDate")) = 0 Or IsNull(rs("LastLogonDate")) Then
			LastLogonDate = "---------"
		Else
			LastLogonDate = rs("LastLogonDate")
		End If
	%>    		
       	<tr>
        		<td valign="top"><font size="2"><a href="_person.asp?PersonID=<%=rs("PersonId")%>"><%=FullName%></a></td>
        		<td valign="top"><font size="2"><%=rs("Organization")%></font></td>
        		<td valign="top"><font size="2"><%=rs("ActiveUser")%></font></td>
        		<td valign="top"><font size="2"><%=rs("Role")%></font></td>
        		<td valign="top"><font size="2"><%=Logon%></font></td>
        		<td valign="top"><font size="2"><%=Phone%></font></td>
        		<td valign="top"><font size="2"><%=Left(LastLogonDate,9)%></font></td>
        	</tr>
	<%	
		rs.movenext
		loop
		rs.close
		set rs = nothing


	%>
   			</table>

</form>
<!--#INCLUDE FILE="_incfooter.asp"-->

</BODY>






















