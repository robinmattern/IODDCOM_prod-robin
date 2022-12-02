<!-- incnav.asp -->
  <center>
  <div align="left">
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
  </center>
    <tr>
      <td align="center"><font size="2" face="Times New Roman">

<%
'==================================================%>

		<a title="Takes you to the home page for <%=session("systemname")%>" href="default.asp">Home</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a title="Takes you to the Logon screen in to gain access to <%=session("systemname")%>.  If you do not have a valid email address, please contact your Point Of Contact in order to establish a valid account" href="_logon.asp">Logon</a>&nbsp;&nbsp;&nbsp;&nbsp;
<% If trim(session("personemail")) <> "" then ' logged in %>
		<a title="Allows to view <%=session("systemname")%> data and if authorized make changes to the data." href="<%=Session("MainTableCall")%>">Data</a>&nbsp;&nbsp;&nbsp;&nbsp;
	<% If false then   'trim(session("personemail")) <> "" then ' not logged in %>
		<a title="Allows you to change the Fiscal Year in <%=session("systemname")%>" href="changefy.asp">FY</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		<a title="Allows you to run reports for <%=session("systemname")%>" href="_rptmenu.asp">Reports</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<% End If%>
	<% If false then   'trim(session("personemail")) <> "" then ' not logged in %>
	    <a title="Mail for <%=session("systemname")%>" href="_email.asp">Mail</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	    <a title="Help information for <%=session("systemname")%>" href="_help.asp">Help</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<% End If%>
	<%If false and Ucase(Session("Role")) = "ADM" Then%>
		<a title="Frequently Asked Questions for <%=session("systemname")%>" href="_faq.asp">FAQ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a title="Score 300.  Provides information about participation in and progress of the data collection effort for <%=session("systemname")%>" href="_notready.asp">Score 300</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<%End If%>
	<% If false and instr(session("grouplist"),"ADM") then %>
	    <a title="Search in <%=session("systemname")%>" href="databasesearch.asp">Search</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<%End If%>
	
	<%If instr(session("grouplist"),"ADM") OR instr(session("grouplist"),"POC") then %>
	    <a title="Takes you to the Administrator screen for <%=session("systemname")%>." href="_adminmenu.asp">Admin</a><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<%End If%>
<%End If%>

</b>

</font>

<b>
<font face="Times New Roman" color="maroon" size="2">
<% If trim(session("personemail")) = "" then ' not logged in %>
	Not Logged On&nbsp;&nbsp;&nbsp;&nbsp
<%Else%>
    <%=session("Personname") & "   (" & session("grouplist") & ")"%>&nbsp;&nbsp;&nbsp;&nbsp
<%End If%>
</font>
</b>
      </td>
    </tr>
  <center>

    <tr>
      <td align="center" height="10">
<font size="1">&nbsp;</font>
      </td>
    </tr>

  </table>
  </div>
  </center>