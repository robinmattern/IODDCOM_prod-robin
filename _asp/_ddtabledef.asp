<% @LANGUAGE="VBSCRIPT" %>
<% Response.Buffer = True %>

<!--#INCLUDE FILE="_inccreateconnection.asp"-->

<%
Set rs = Conn.Execute("SELECT * FROM tfmssystem")
%>

<HTML>
<HEAD>
<TITLE>Table Definition- tfmssystems</TITLE>
</HEAD>
<BODY>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->

&nbsp;
<div align="center">
  <center>

<TABLE BORDER=3 CELLPADING=0 CELLSPACING=0 cellpadding="2" style="border-collapse: collapse" bordercolor="#000080">
<TR>
	<TD bgcolor="#99CCFF" colspan="4">
    <p align="center"><b><font size="4">Table Definition- tfmssystems</font></b></TD>
	</TR>
<TR>
	<TD bgcolor="#99CCFF"><b>Field Name</b></TD>
	<TD bgcolor="#99CCFF"><b>Field Type</b></TD>
	<TD bgcolor="#99CCFF"><b>Field Size</b></TD>
	<TD bgcolor="#99CCFF"><b>Sample Data</b></TD>
</TR>
<%
For Each oFld in rs.Fields
%>
	<TR>
		<TD><%=oFld.Name%>&nbsp;</TD>
		<%	Select Case oFld.Type
			Case "3"
				fldType = "int"
			Case "129"
				fldType = "char"
			Case "131"
				fldType = "numeric"
			Case "135"
				fldType = "datetime"
			Case "200"
				fldType = "varchar"
			Case Else
				fldType = oFld.Type
			End Select	
		%>
		<TD><%=FldType%>&nbsp;</TD>
		<TD><%=oFld.DefinedSize%>&nbsp;</TD>
		<%	'If oFld.Properties > 0 then %>
		<%	'End if %>
		<TD><%=oFld.Value%>&nbsp;</TD>
	</TR>	
<%
Next

rs.Close
Set rs = Nothing
Conn.Close
Set Conn = Nothing
%>

</TABLE>
  </center>
</div>
</BODY>
<!--#INCLUDE FILE="_incfooter.asp"-->
</HTML>