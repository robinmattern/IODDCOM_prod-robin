<%@ LANGUAGE="VBScript" %>

<%
'-------------------------------------------------------------------------------
' Microsoft Visual InterDev - Data Form Wizard
' 
' List Page
'
' (c) 1997 Microsoft Corporation.  All Rights Reserved.
'
' This file is an Active Server Page that contains the list view of a Data Form. 
' It requires Microsoft Internet Information Server 3.0 and can be displayed
' using any browser that supports tables. You can edit this file to further 
' customize the list view.
'
'-------------------------------------------------------------------------------

Dim strPagingMove
Dim strDFName
strDFName = "rsPantPan"
%>

<script RUNAT="Server" LANGUAGE="VBScript">

'---- DataTypeEnum Values ----
Const adUnsignedTinyInt = 17
Const adBoolean = 11
Const adLongVarChar = 201
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205

'-------------------------------------------------------------------------------
' Purpose:  Substitutes Empty for Null and trims leading/trailing spaces
' Inputs:   varTemp	- the target value
' Returns:	The processed value
'-------------------------------------------------------------------------------

Function ConvertNull(varTemp)
	If IsNull(varTemp) Then
		ConvertNull = ""
	Else
		ConvertNull = Trim(varTemp)
	End If
End Function

'-------------------------------------------------------------------------------
' Purpose:  Embeds bracketing quotes around the string
' Inputs:   varTemp	- the target value
' Returns:	The processed value
'-------------------------------------------------------------------------------

Function QuotedString(varTemp)
	If IsNull(varTemp) Then
		QuotedString = Chr(34) & Chr(34)
	Else
		QuotedString = Chr(34) & CStr(varTemp) & Chr(34)
	End If
End Function

'-------------------------------------------------------------------------------
' Purpose:  Tests string to see if it is a URL by looking for protocol
' Inputs:   varTemp	- the target value
' Returns:	True - if is URL, False if not
'-------------------------------------------------------------------------------

Function IsURL(varTemp)
	IsURL = True
	If UCase(Left(Trim(varTemp), 6)) = "HTTP:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 6)) = "FILE:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 8)) = "MAILTO:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 5)) = "FTP:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 8)) = "GOPHER:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 6)) = "NEWS:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 7)) = "HTTPS:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 8)) = "TELNET:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 6)) = "NNTP:/" Then Exit Function
	IsURL = False
End Function

'-------------------------------------------------------------------------------
' Purpose:  Handles the display of a field from a recordset depending
'			on its data type, attributes, and the current mode.
' Assumes: 	That the recordset containing the field is open
' Inputs:   strFieldName 	- the name of the field in the recordset
'			avarLookup		- array of lookup values
'-------------------------------------------------------------------------------
 
Function ShowField(strFieldName, avarLookup)
	Dim intRow
	Dim strPartial
	Dim strBool
	Dim nPos
	strFieldValue = ""
	nPos=Instr(strFieldName,".")
	Do While nPos > 0 
		strFieldName= Mid (strFieldName, nPos+1)
		nPos=Instr(strFieldName,".")
	Loop 
	If Not IsNull(avarLookup) Then
		Response.Write "<TD BGCOLOR=White NOWRAP><FONT SIZE=-1>" 
		For intRow = 0 to UBound(avarLookup, 2)
			If ConvertNull(avarLookup(0, intRow)) = ConvertNull(rsPantPan(strFieldName)) Then
				Response.Write Server.HTMLEncode(ConvertNull(avarLookup(1, intRow)))
				Exit For
			End If
		Next
		Response.Write "</FONT></TD>"
		Exit Function
	End If
	
	Select Case rsPantPan(strFieldName).Type
		Case adBoolean, adUnsignedTinyInt				'Boolean
			strBool = ""
			If rsPantPan(strFieldName) <> 0 Then
				strBool = "True"
			Else
				strBool = "False"
			End If
			Response.Write "<TD BGCOLOR=White ><FONT SIZE=-1>" & strBool & "</FONT></TD>"
			
		Case adBinary, adVarBinary, adLongVarBinary		'Binary
			Response.Write "<TD BGCOLOR=White ><FONT SIZE=-1>[Binary]</FONT></TD>"
			
		Case adLongVarChar, adLongVarWChar				'Memo
			Response.Write "<TD BGCOLOR=White NOWRAP><FONT SIZE=-1>"
			strPartial = Left(rsPantPan(strFieldName), 50)			
			If ConvertNull(strPartial) = "" Then
				Response.Write "&nbsp;"
			Else
				Response.Write Server.HTMLEncode(strPartial)
			End If
			If rsPantPan(strFieldName).ActualSize > 50 Then Response.Write "..."
			Response.Write "</FONT></TD>"
			
		Case Else
			Response.Write "<TD BGCOLOR=White ALIGN=Left NOWRAP><FONT SIZE=-1>"
			If ConvertNull(rsPantPan(strFieldName)) = "" Then
				Response.Write "&nbsp;"
			Else
				' Check for special field types
				Select Case UCase(Left(rsPantPan(strFieldName).Name, 4))
					Case "URL_"
						Response.Write "<A HREF=" & QuotedString(rsPantPan(strFieldName)) & ">"
						Response.Write Server.HTMLEncode(ConvertNull(rsPantPan(strFieldName)))
						Response.Write "</A>"
					Case Else
						If IsURL(rsPantPan(strFieldName)) Then
							Response.Write "<A HREF=" & QuotedString(rsPantPan(strFieldName)) & ">"
							Response.Write Server.HTMLEncode(ConvertNull(rsPantPan(strFieldName)))
							Response.Write "</A>"
						Else
							Response.Write Server.HTMLEncode(ConvertNull(rsPantPan(strFieldName)))
						End If
				End Select
			End If
			Response.Write "</FONT></TD>"
	End Select
End Function

</script>

<html>
<head>
	<meta NAME="GENERATOR" CONTENT="Microsoft Visual InterDev">
	
	<meta NAME="Keywords" CONTENT="Microsoft Data Form, tPan List">
	<title>tPan List</title>
<% ' FP_ASP -- ASP Automatically generated by a Frontpage Component. Do not Edit.
FP_CharSet = "windows-1252"
FP_CodePage = 1252 %>
</head>

<!--------------------------- Formatting Section ------------------------------>

<basefont FACE="Arial, Helvetica, sans-serif">
<body>

<!---------------------------- Lookups Section ------------------------------->
<% 
Dim avarPantMember
If IsEmpty(Application(strDFName & "_Lookup_PantMember")) Or strPagingMove = "Requery" Then
    Set DataConn = Server.CreateObject("ADODB.Connection")
    DataConn.ConnectionTimeout = Session("DataConn_ConnectionTimeout")
    DataConn.CommandTimeout = Session("DataConn_CommandTimeout")
    DataConn.Open Session("DataConn_ConnectionString"), Session("DataConn_RuntimeUserName"), Session("DataConn_RuntimePassword")
	Set rsPantMember = DataConn.Execute("SELECT DISTINCT `MemberID`, `LastName` FROM `tMember`")
	avarPantMember = Null
	On Error Resume Next
	avarPantMember = rsPantMember.GetRows()
	Application.Lock
	Application(strDFName & "_Lookup_PantMember") = avarPantMember
	Application.Unlock
Else
	avarPantMember = Application(strDFName & "_Lookup_PantMember")
End If
%>


<!---------------------------- Heading Section ------------------------------->

<% Response.Write "<FORM ACTION=PanForm.asp METHOD=""POST"">" %>
<table WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
	<tr>
		<th NOWRAP BGCOLOR="green" ALIGN="Left">
			<font SIZE="5">&nbsp;Pans</font>
		</th>
		<td BGCOLOR="green" VALIGN="Middle" ALIGN="Right" WIDTH="100%">
			<input TYPE="Hidden" NAME="FormMode" VALUE="Edit">
			&nbsp;<input TYPE="SUBMIT" NAME="DataAction" VALUE="Form View">&nbsp;
		</td>
	</tr>
	<tr>
		<td BGCOLOR="#FFFFCC" COLSPAN="3">
			<font SIZE="-1">&nbsp;&nbsp;
			<% 
			If IsEmpty(Session("rsPantPan_Filter")) Or Session("rsPantPan_Filter")="" Then
				Response.Write "Current Filter: None<BR>"
			Else
				Response.Write "Current Filter: " & Session("rsPantPan_FilterDisplay") & "<BR>"
			End If 
			%>
            </font>
        </td>
    </tr></table>
</form>

<!----------------------------- List Section --------------------------------->

<table CELLSPACING="0" CELLPADDING="0" BORDER="0" WIDTH="100%">
<tr>
<td WIDTH="20">&nbsp;</td>
<td>
<table CELLSPACING="1" CELLPADDING="1" BORDER="0" WIDTH="100%">
	<tr BGCOLOR="green" VALIGN="TOP">
		<td ALIGN="Center"><font SIZE="-1">&nbsp;#&nbsp;</font></td>
		<td ALIGN="Left"><font SIZE="-1"><b>PanID</b></font></td>
		<td ALIGN="Left"><font SIZE="-1"><b>MemberID</b></font></td>
		<td ALIGN="Left"><font SIZE="-1"><b>Pan</b></font></td>
		<td ALIGN="Left"><font SIZE="-1"><b>Comment</b></font></td>
		<td ALIGN="Left"><font SIZE="-1"><b>Date</b></font></td>
	</tr>

<!--METADATA TYPE="DesignerControl" startspan
	<OBJECT ID="rsPantPan" WIDTH=151 HEIGHT=24
		CLASSID="CLSID:F602E721-A281-11CF-A5B7-0080C73AAC7E">
		<PARAM NAME="BarAlignment" VALUE="0">
       	<PARAM NAME="PageSize" VALUE="10">
		<PARAM Name="RangeType" Value="2">
		<PARAM Name="DataConnection" Value="DataConn">
		<PARAM Name="CommandType" Value="0">
		<PARAM Name="CommandText" Value="SELECT `PanID`, `MemberID`, `Pan`, `Comment`, `Date` FROM `tPan`">
		<PARAM Name="CursorType" Value="1">
		<PARAM Name="LockType" Value="3">
		<PARAM Name="CacheRecordset" Value="1">
    </OBJECT>
-->

<%
fHideNavBar = False
fHideNumber = False
fHideRequery = False
fHideRule = False
stQueryString = ""
fEmptyRecordset = False
fFirstPass = True
fNeedRecordset = False
fNoRecordset = False
tBarAlignment = "Left"
tHeaderName = "rsPantPan"
tPageSize = 10
tPagingMove = ""
tRangeType = "Table"
tRecordsProcessed = 0
tPrevAbsolutePage = 0
intCurPos = 0
intNewPos = 0
fSupportsBookmarks = True
fMoveAbsolute = False

If Not IsEmpty(Request("rsPantPan_PagingMove")) Then
    tPagingMove = Trim(Request("rsPantPan_PagingMove"))
End If

If IsEmpty(Session("rsPantPan_Recordset")) Then
    fNeedRecordset = True
Else
    If Session("rsPantPan_Recordset") Is Nothing Then
        fNeedRecordset = True
    Else
        Set rsPantPan = Session("rsPantPan_Recordset")
    End If
End If

If fNeedRecordset Then
    Set DataConn = Server.CreateObject("ADODB.Connection")
    DataConn.ConnectionTimeout = Session("DataConn_ConnectionTimeout")
    DataConn.CommandTimeout = Session("DataConn_CommandTimeout")
    DataConn.Open Session("DataConn_ConnectionString"), Session("DataConn_RuntimeUserName"), Session("DataConn_RuntimePassword")
    Set cmdTemp = Server.CreateObject("ADODB.Command")
    Set rsPantPan = Server.CreateObject("ADODB.Recordset")
    cmdTemp.CommandText = "SELECT `PanID`, `MemberID`, `Pan`, `Comment`, `Date` FROM `tPan`"
    cmdTemp.CommandType = 1
    Set cmdTemp.ActiveConnection = DataConn
    rsPantPan.Open cmdTemp, , 1, 3
End If
On Error Resume Next
If rsPantPan.BOF And rsPantPan.EOF Then fEmptyRecordset = True
On Error Goto 0
If Err Then fEmptyRecordset = True
If fNeedRecordset Then
    Set Session("rsPantPan_Recordset") = rsPantPan
End If
rsPantPan.PageSize = tPageSize
fSupportsBookmarks = rsPantPan.Supports(8192)

If Not IsEmpty(Session("rsPantPan_Filter")) And Not fEmptyRecordset Then
    rsPantPan.Filter = Session("rsPantPan_Filter")
    If rsPantPan.BOF And rsPantPan.EOF Then fEmptyRecordset = True
End If

If IsEmpty(Session("rsPantPan_PageSize")) Then Session("rsPantPan_PageSize") = tPageSize
If IsEmpty(Session("rsPantPan_AbsolutePage")) Then Session("rsPantPan_AbsolutePage") = 1

If Session("rsPantPan_PageSize") <> tPageSize Then
    tCurRec = ((Session("rsPantPan_AbsolutePage") - 1) * Session("rsPantPan_PageSize")) + 1
    tNewPage = Int(tCurRec / tPageSize)
    If tCurRec Mod tPageSize <> 0 Then
        tNewPage = tNewPage + 1
    End If
    If tNewPage = 0 Then tNewPage = 1
    Session("rsPantPan_PageSize") = tPageSize
    Session("rsPantPan_AbsolutePage") = tNewPage
End If

If fEmptyRecordset Then
    fHideNavBar = True
    fHideRule = True
Else
    tPrevAbsolutePage = Session("rsPantPan_AbsolutePage")
    Select Case tPagingMove
        Case ""
            fMoveAbsolute = True
        Case "Requery"
            rsPantPan.Requery
            fMoveAbsolute = True
        Case "<<"
            Session("rsPantPan_AbsolutePage") = 1
        Case "<"
            If Session("rsPantPan_AbsolutePage") > 1 Then
                Session("rsPantPan_AbsolutePage") = Session("rsPantPan_AbsolutePage") - 1
            End If
        Case ">"
            If Not rsPantPan.EOF Then
                Session("rsPantPan_AbsolutePage") = Session("rsPantPan_AbsolutePage") + 1
            End If
        Case ">>"
            If fSupportsBookmarks Then
                Session("rsPantPan_AbsolutePage") = rsPantPan.PageCount
            End If
    End Select
    Do
        If fSupportsBookmarks Then
            rsPantPan.AbsolutePage = Session("rsPantPan_AbsolutePage")
        Else
            If fNeedRecordset Or fMoveAbsolute Or rsPantPan.EOF Or Not fSupportsMovePrevious Then
                rsPantPan.MoveFirst
                rsPantPan.Move (Session("rsPantPan_AbsolutePage") - 1) * tPageSize
            Else
                intCurPos = ((tPrevAbsolutePage - 1) * tPageSize) + tPageSize
                intNewPos = ((Session("rsPantPan_AbsolutePage") - 1) * tPageSize) + 1
                rsPantPan.Move intNewPos - intCurPos
            End If
            If rsPantPan.BOF Then rsPantPan.MoveNext
        End If
        If Not rsPantPan.EOF Then Exit Do
        Session("rsPantPan_AbsolutePage") = Session("rsPantPan_AbsolutePage") - 1
    Loop
End If

Do
    If fEmptyRecordset Then Exit Do
    If tRecordsProcessed = tPageSize Then Exit Do
    If Not fFirstPass Then
        rsPantPan.MoveNext
    Else
        fFirstPass = False
    End If
    If rsPantPan.EOF Then Exit Do
    tRecordsProcessed = tRecordsProcessed + 1
%>
<!--METADATA TYPE="DesignerControl" endspan-->

	<tr VALIGN="TOP">
		<td BGCOLOR="White"><font SIZE="-1">
		<%
		If tPageSize > 0 Then
			tCurRec = ((Session("rsPantPan_AbsolutePage") - 1) * tPageSize) + tRecordsProcessed
		Else
			tRecordsProcessed = tRecordsProcessed + 1
			tCurRec = tRecordsProcessed
		End If
		Response.Write "<A HREF=" & QuotedString("PanAction.asp?Bookmark=" & tCurRec & "&DataAction=Find") & ">" & tCurRec & "</A>"
		%>

		</font></td>
		<%
		ShowField "PanID", Null
		ShowField "MemberID",avarPantMember
		ShowField "Pan", Null
		ShowField "Comment", Null
		ShowField "Date", Null
		fHideRule = True
		%>
	</tr>

<!--METADATA TYPE="DesignerControl" startspan
    <OBJECT ID="DataRangeFtr1" WIDTH=151 HEIGHT=24
        CLASSID="CLSID:F602E722-A281-11CF-A5B7-0080C73AAC7E">
    </OBJECT>
-->
<%
Loop
If tRangeType = "Table" Then Response.Write "</TABLE>"
If tPageSize > 0 Then
    If Not fHideRule Then Response.Write "<HR>"
    If Not fHideNavBar Then
        %>
        <TABLE WIDTH=100% >
        <TR>
            <TD WIDTH=100% >
                <P ALIGN=<%= tBarAlignment %> >
                <FORM <%= "ACTION=""" & Request.ServerVariables("PATH_INFO") & stQueryString & """" %> METHOD="POST">
                    <INPUT TYPE="Submit" NAME="<%= tHeaderName & "_PagingMove" %>" VALUE="   &lt;&lt;   ">
                    <INPUT TYPE="Submit" NAME="<%= tHeaderName & "_PagingMove" %>" VALUE="   &lt;    ">
                    <INPUT TYPE="Submit" NAME="<%= tHeaderName & "_PagingMove" %>" VALUE="    &gt;   ">
                    <% If fSupportsBookmarks Then %>
                        <INPUT TYPE="Submit" NAME="<%= tHeaderName & "_PagingMove" %>" VALUE="   &gt;&gt;   ">
                    <% End If %>
                    <% If Not fHideRequery Then %>
                        <INPUT TYPE="Submit" NAME="<% =tHeaderName & "_PagingMove" %>" VALUE=" Requery ">
                    <% End If %>
                </FORM>
                </P>
            </TD>
            <TD VALIGN=MIDDLE ALIGN=RIGHT>
                <FONT SIZE=2>
                <%
                If Not fHideNumber Then
                    If tPageSize > 1 Then
                        Response.Write "<NOBR>Page: " & Session(tHeaderName & "_AbsolutePage") & "</NOBR>"
                    Else
                        Response.Write "<NOBR>Record: " & Session(tHeaderName & "_AbsolutePage") & "</NOBR>"
                    End If
                End If
                %>
                </FONT>
            </TD>
        </TR>
        </TABLE>
    <%
    End If
End If
%>
<!--METADATA TYPE="DesignerControl" endspan-->

<!---------------------------- Footer Section -------------------------------->

<% 
' TEMP: cache here until CacheRecordset property is implemented in
' 		data range
If fNeedRecordset Then
	Set Session("rsPantPan_Recordset") = rsPantPan
End If 
%>

</td></tr></table>
</body>
</html>

