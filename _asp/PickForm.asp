<%@ LANGUAGE="vbscript" %>
<%
'-------------------------------------------------------------------------------
' Microsoft Visual InterDev - Data Form Wizard
' 
' Form Page
'
' (c) 1997 Microsoft Corporation.  All Rights Reserved.
'
' This file is an Active Server Page that contains the form view of a Data Form. 
' It requires Microsoft Internet Information Server 3.0 and can be displayed
' using any browser that supports tables. You can edit this file to further 
' customize the form view.
'
' Modes: 		The form mode can be controlled by passing the following
'				name/value pairs using POST or GET:
'				FormMode=Edit
'				FormMode=Filter
'				FormMode=New
' Tips:			- If a field contains a URL to an image and has a name that 
'				begins with "img_" (case-insensitive), the image will be 
'				displayed using the IMG tag.
'				- If a field contains a URL and has a name that begins with 
'				"url_" (case-insensitive), a jump will be displayed using the 
'				Anchor tag.
'-------------------------------------------------------------------------------

Dim strPagingMove	
Dim strFormMode
Dim strDFName
strDFName = "rsPicktPick"
%>

<script RUNAT="Server" LANGUAGE="VBScript">

'---- FieldAttributeEnum Values ----
Const adFldUpdatable = &H00000004
Const adFldUnknownUpdatable = &H00000008
Const adFldIsNullable = &H00000020

'---- DataTypeEnum Values ----
Const adUnsignedTinyInt = 17
Const adBoolean = 11
Const adLongVarChar = 201
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205
Const adVarChar = 200
Const adWVarChar = 202
Const adBSTR = 8
Const adChar = 129
Const adWChar = 130
'---- Other Values ----
Const dfMaxSize = 100

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
' Purpose:  Tests whether the field in the recordset is updatable
' Assumes: 	That the recordset containing the field is open
' Inputs:   strFieldName	- the name of the field in the recordset
' Returns:	True if updatable, False if not
'-------------------------------------------------------------------------------

Function CanUpdateField(strFieldName)
	Dim intUpdatable
	intUpdatable = (adFldUpdatable Or adFldUnknownUpdatable)
	CanUpdateField = True
	If (rsPicktPick(strFieldName).Attributes And intUpdatable) = False Then
		CanUpdateField = False
	End If
End Function

'-------------------------------------------------------------------------------
' Purpose:  Handles the display of a field from a recordset depending
'			on its data type, attributes, and the current mode.
' Assumes: 	That the recordset containing the field is open
'			That strFormMode is initialized
' Inputs:   strFieldName 	- the name of the field in the recordset
'			strLabel		- the label to display
'			blnIdentity		- identity field flag
'			avarLookup		- array of lookup values
'-------------------------------------------------------------------------------
 
Sub ShowField(strFieldName, strLabel, blnIdentity, avarLookup)
	Dim blnFieldRequired
	Dim intMaxSize
	Dim intInputSize
	Dim strOption1State
	Dim strOption2State
	Dim strFieldValue
	Dim nPos
	strFieldValue = ""
	nPos=Instr(strFieldName,".")
	Do While nPos > 0 
		strFieldName= Mid (strFieldName, nPos+1)
		nPos=Instr(strFieldName,".")
	Loop 
	' If not in Edit form mode then set value to empty so doesn't display
	strFieldValue = ""
	If strFormMode = "Edit" Then strFieldValue = RTrim(rsPicktPick(strFieldName))
	
	' See if the field is required by checking the attributes 
	blnFieldRequired = False
	If (rsPicktPick(strFieldName).Attributes And adFldIsNullable) = 0 Then 
		blnFieldRequired = True
	End If
	
	' Set values for the MaxLength and Size attributes	
	intMaxSize = dfMaxSize
	intInputSize = rsPicktPick(strFieldName).DefinedSize + 2
	If strFormMode <> "Filter" Then intMaxSize = intInputSize - 2
	
	' Write the field label and start the value cell
	Response.Write "<TR VALIGN=TOP>"
	Response.Write "<TD HEIGHT=25 ALIGN=Left NOWRAP><FONT SIZE=-1><B>&nbsp;&nbsp;" & strLabel & "</B></FONT></TD>"	
	Response.Write "<TD WIDTH=100% ><FONT SIZE=-1>"
	
	' If the field is not updatable, then handle 
	' it like an Identity column and exit
	If Not CanUpdateField(strFieldName) Then
		' Special handling if Binary
		Select Case rsPicktPick(strFieldName).Type
			Case adBinary, adVarBinary, adLongVarBinary		'Binary
				Response.Write "[Binary]"
			Case Else		
				Select Case strFormMode
					Case "Edit"
						Response.Write ConvertNull(strFieldValue)
						Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strFieldName)
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
					Case "New"
						Response.Write "[AutoNumber]"
						Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strFieldName)
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
					Case "Filter" 
						Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strFieldName)
						Response.Write " SIZE=" & intInputSize
						Response.Write " MAXLENGTH=" & intMaxSize
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
				End Select
		End Select
		Response.Write "</FONT></TD></TR>"
		Exit Sub
	End If
	
	' Handle lookups using a select and options
	If Not IsNull(avarLookup) Then
		Response.Write "<SELECT NAME=" & QuotedString(strFieldName) & ">"
		' Add blank entry if not required or in filter mode
		If Not blnFieldRequired Or strFormMode = "Filter" Then
			If (strFormMode = "Filter" Or strFormMode = "New") Then
				Response.Write "<OPTION SELECTED>"
			Else
				Response.Write "<OPTION>"
			End If
		End If
		
		' Loop thru the rows in the array
		For intRow = 0 to UBound(avarLookup, 2)
			Response.Write "<OPTION VALUE=" & QuotedString(avarLookup(0, intRow))
            If strFormMode = "Edit" Then
				If ConvertNull(avarLookup(0, intRow)) = ConvertNull(strFieldValue) Then
               		Response.Write " SELECTED"
				End If
            End If
           	Response.Write ">"
			Response.Write ConvertNull(avarLookup(1, intRow))
		Next
		Response.Write "</SELECT>"
		If blnFieldRequired And strFormMode = "New" Then 
			Response.Write "  Required"
		End If
		Response.Write "</FONT></TD></TR>"
		Exit Sub
	End If	
	
	' Evaluate data type and handle appropriately
	Select Case rsPicktPick(strFieldName).Type
	
		Case adBoolean, adUnsignedTinyInt				'Boolean
			If strFormMode = "Filter" Then				
				strOption1State = " >True"
				strOption2State = " >False"
			Else
				Select Case strFieldValue
					Case "True", "1", "-1"
						strOption1State = " CHECKED>True"
						strOption2State = " >False"
					Case "False", "0"
						strOption1State = " >True"
						strOption2State = " CHECKED>False"
					Case Else
						strOption1State = " >True"
						strOption2State = " >False"
				End Select
			End If			
			Response.Write "<INPUT TYPE=Radio VALUE=1 NAME=" & QuotedString(strFieldName) & strOption1State
			Response.Write "<INPUT TYPE=Radio VALUE=0 NAME=" & QuotedString(strFieldName) & strOption2State
			If strFormMode = "Filter" Then
				Response.Write "<INPUT TYPE=Radio NAME=" & QuotedString(strFieldName) & " CHECKED>Neither"
			End If
			
		Case adBinary, adVarBinary, adLongVarBinary		'Binary
			Response.Write "[Binary]"
			
		Case adLongVarChar, adLongVarWChar				'Memo
			Response.Write "<TEXTAREA NAME=" & QuotedString(strFieldName) & " ROWS=3 COLS=80>"
			Response.Write Server.HTMLEncode(ConvertNull(strFieldValue))
			Response.Write "</TEXTAREA>"
			
		Case Else
			Dim nType 
			nType=rsPicktPick(strFieldName).Type
			If (nType <> adVarChar) and (nType <> adWVarChar) and (nType <> adBSTR) and (nType <> adChar) and (nType <> adWChar)  Then
				intInputSize = (intInputSize-2)*3+2
				If strFormMode <> "Filter" Then intMaxSize = intInputSize - 2
			End If
			If blnIdentity Then
				Select Case strFormMode
					Case "Edit"
						Response.Write ConvertNull(strFieldValue)
						Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strFieldName)
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
					Case "New"
						Response.Write "[AutoNumber]"
						Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strFieldName)
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
					Case "Filter" 
						Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strFieldName) & " SIZE=" & tInputSize
						Response.Write " MAXLENGTH=" & tMaxSize & " VALUE=" & QuotedString(strFieldValue) & " >"
				End Select
			Else
				If intInputSize > 80 Then intInputSize = 80			
				Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strFieldName)
				Response.Write " SIZE=" & intInputSize
				Response.Write " MAXLENGTH=" & intMaxSize
				Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
				' Check for special field types
				Select Case UCase(Left(rsPicktPick(strFieldName).Name, 4))
					Case "IMG_"
						If strFieldValue <> "" Then
							Response.Write "<BR><BR><IMG SRC=" & QuotedString(strFieldValue) & "><BR>&nbsp;<BR>"
						End If
					Case "URL_"
						If strFieldValue <> "" Then
							Response.Write "&nbsp;&nbsp;<A HREF=" & QuotedString(strFieldValue) & ">"
							Response.Write "Go"
							Response.Write "</A>"
						End If
					Case Else
						If IsURL(strFieldValue) Then
							Response.Write "&nbsp;&nbsp;<A HREF=" & QuotedString(strFieldValue) & ">"
							Response.Write "Go"
							Response.Write "</A>"
						End If					
				End Select				
			End If
	End Select
   	If blnFieldRequired And strFormMode = "New" Then
		Response.Write "  Required"
	End If
	Response.Write "</FONT></TD></TR>"
End Sub	
</script>

<% 
strFormMode = "Edit"	' Initalize the default
If Not IsEmpty(Request("FormMode")) Then strFormMode = Request("FormMode")
If Not IsEmpty(Request("rsPicktPick_PagingMove")) Then
    strPagingMove = Trim(Request("rsPicktPick_PagingMove"))
End If
%>

<html>
<head>
	<meta NAME="GENERATOR" CONTENT="Microsoft Visual InterDev">
	
	<meta NAME="Keywords" CONTENT="Microsoft Data Form, tPick Form">
	<title>tPick Form</title>
<% ' FP_ASP -- ASP Automatically generated by a Frontpage Component. Do not Edit.
FP_CharSet = "windows-1252"
FP_CodePage = 1252 %>
</head>

<!--------------------------- Formatting Section ------------------------------>

<basefont FACE="Arial, Helvetica, sans-serif">
<body>

<!---------------------------- Lookups Section -------------------------------->
<% 
Dim avarPicktMember
If IsEmpty(Application(strDFName & "_Lookup_PicktMember")) Or strPagingMove = "Requery" Then
    Set DataConn = Server.CreateObject("ADODB.Connection")
    DataConn.ConnectionTimeout = Session("DataConn_ConnectionTimeout")
    DataConn.CommandTimeout = Session("DataConn_CommandTimeout")
    DataConn.Open Session("DataConn_ConnectionString"), Session("DataConn_RuntimeUserName"), Session("DataConn_RuntimePassword")
	Set rsPicktMember = DataConn.Execute("SELECT DISTINCT `MemberID`, `LastName` FROM `tMember`")
	avarPicktMember = Null
	On Error Resume Next
	avarPicktMember = rsPicktMember.GetRows()
	Application.Lock
	Application(strDFName & "_Lookup_PicktMember") = avarPicktMember
	Application.Unlock
Else
	avarPicktMember = Application(strDFName & "_Lookup_PicktMember")
End If
%>


<!---------------------------- Heading Section -------------------------------->

<% Response.Write "<FORM ACTION=""PickAction.asp"" METHOD=""POST"">" %>
<table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
		<th NOWRAP BGCOLOR="Green">
			<font SIZE="5">&nbsp;Picks&nbsp;</font>
		</th>
		<td ALIGN="Right" BGCOLOR="Green" VALIGN="MIDDLE" WIDTH="100%">
			<% 
			If strFormMode = "Form View" then strFormMode = "Edit"
			Select Case strFormMode
				Case "Edit"	
					%>
					<input TYPE="SUBMIT" NAME="DataAction" VALUE="Update">
					<input TYPE="SUBMIT" NAME="DataAction" VALUE="Delete">
					<input TYPE="SUBMIT" NAME="DataAction" VALUE="New">
					<input TYPE="SUBMIT" NAME="DataAction" VALUE="Filter">
					<% If Session("rsPicktPick_Filter") <> "" Then %>
						&nbsp;&nbsp;<input TYPE="SUBMIT" NAME="DataAction" VALUE="All Records">
					<% End If %>&nbsp;
				<% Case "Filter" %>
					<input TYPE="SUBMIT" NAME="DataAction" VALUE=" Apply ">
					<input TYPE="SUBMIT" NAME="DataAction" VALUE=" Cancel ">&nbsp;
				<% Case "New" %>
					<input TYPE="SUBMIT" NAME="DataAction" VALUE=" Insert ">
					<input TYPE="SUBMIT" NAME="DataAction" VALUE=" Cancel ">&nbsp;
			<% End Select %>
			&nbsp;<input TYPE="SUBMIT" NAME="DataAction" VALUE="List View">&nbsp;		
		</td>
    </tr>
	<tr>
		<td BGCOLOR="#FFFFCC" COLSPAN="3">
			<font SIZE="-1">&nbsp;&nbsp;
			<%
			If Not IsEmpty(Session("rsPicktPick_Status")) And Session("rsPicktPick_Status") <>"" Then
				Response.Write Session("rsPicktPick_Status")
				Session("rsPicktPick_Status") = ""
			Else
				Select Case strFormMode
					Case "Edit"
						If IsEmpty(Session("rsPicktPick_Filter")) Then
							Response.Write "Current Filter: None"
						Else
							If Session("rsPicktPick_Filter") <> "" Then
								Response.Write "Current Filter: " & Session("rsPicktPick_FilterDisplay")
							Else
								Response.Write "Current Filter: None"
							End If
						End If
					Case "Filter"
						Response.Write "Status: Ready for filter criteria"
					Case "New"
						Response.Write "Status: Ready for new record"
				End Select
			End If 
			%>
			</font>
		</td>
	</tr></table>

<!----------------------------- Form Section ---------------------------------->

<!--METADATA TYPE="DesignerControl" startspan
	<OBJECT ID="rsPicktPick" WIDTH=151 HEIGHT=24
		CLASSID="CLSID:F602E721-A281-11CF-A5B7-0080C73AAC7E">
		<PARAM NAME="BarAlignment" VALUE="0">
       	<PARAM NAME="PageSize" VALUE="1">
		<PARAM Name="RangeType" Value="1">
		<PARAM Name="DataConnection" Value="DataConn">
		<PARAM Name="CommandType" Value="0">
		<PARAM Name="CommandText" Value="SELECT `PickID`, `MemberID`, `Pick`, `Comment`, `Date` FROM `tPick`">
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
tHeaderName = "rsPicktPick"
tPageSize = 1
tPagingMove = ""
tRangeType = "Form"
tRecordsProcessed = 0
tPrevAbsolutePage = 0
intCurPos = 0
intNewPos = 0
fSupportsBookmarks = True
fMoveAbsolute = False

If Not IsEmpty(Request("rsPicktPick_PagingMove")) Then
    tPagingMove = Trim(Request("rsPicktPick_PagingMove"))
End If

If IsEmpty(Session("rsPicktPick_Recordset")) Then
    fNeedRecordset = True
Else
    If Session("rsPicktPick_Recordset") Is Nothing Then
        fNeedRecordset = True
    Else
        Set rsPicktPick = Session("rsPicktPick_Recordset")
    End If
End If

If fNeedRecordset Then
    Set DataConn = Server.CreateObject("ADODB.Connection")
    DataConn.ConnectionTimeout = Session("DataConn_ConnectionTimeout")
    DataConn.CommandTimeout = Session("DataConn_CommandTimeout")
    DataConn.Open Session("DataConn_ConnectionString"), Session("DataConn_RuntimeUserName"), Session("DataConn_RuntimePassword")
    Set cmdTemp = Server.CreateObject("ADODB.Command")
    Set rsPicktPick = Server.CreateObject("ADODB.Recordset")
    cmdTemp.CommandText = "SELECT `PickID`, `MemberID`, `Pick`, `Comment`, `Date` FROM `tPick`"
    cmdTemp.CommandType = 1
    Set cmdTemp.ActiveConnection = DataConn
    rsPicktPick.Open cmdTemp, , 1, 3
End If
On Error Resume Next
If rsPicktPick.BOF And rsPicktPick.EOF Then fEmptyRecordset = True
On Error Goto 0
If Err Then fEmptyRecordset = True
If fNeedRecordset Then
    Set Session("rsPicktPick_Recordset") = rsPicktPick
End If
rsPicktPick.PageSize = tPageSize
fSupportsBookmarks = rsPicktPick.Supports(8192)

If Not IsEmpty(Session("rsPicktPick_Filter")) And Not fEmptyRecordset Then
    rsPicktPick.Filter = Session("rsPicktPick_Filter")
    If rsPicktPick.BOF And rsPicktPick.EOF Then fEmptyRecordset = True
End If

If IsEmpty(Session("rsPicktPick_PageSize")) Then Session("rsPicktPick_PageSize") = tPageSize
If IsEmpty(Session("rsPicktPick_AbsolutePage")) Then Session("rsPicktPick_AbsolutePage") = 1

If Session("rsPicktPick_PageSize") <> tPageSize Then
    tCurRec = ((Session("rsPicktPick_AbsolutePage") - 1) * Session("rsPicktPick_PageSize")) + 1
    tNewPage = Int(tCurRec / tPageSize)
    If tCurRec Mod tPageSize <> 0 Then
        tNewPage = tNewPage + 1
    End If
    If tNewPage = 0 Then tNewPage = 1
    Session("rsPicktPick_PageSize") = tPageSize
    Session("rsPicktPick_AbsolutePage") = tNewPage
End If

If fEmptyRecordset Then
    fHideNavBar = True
    fHideRule = True
Else
    tPrevAbsolutePage = Session("rsPicktPick_AbsolutePage")
    Select Case tPagingMove
        Case ""
            fMoveAbsolute = True
        Case "Requery"
            rsPicktPick.Requery
            fMoveAbsolute = True
        Case "<<"
            Session("rsPicktPick_AbsolutePage") = 1
        Case "<"
            If Session("rsPicktPick_AbsolutePage") > 1 Then
                Session("rsPicktPick_AbsolutePage") = Session("rsPicktPick_AbsolutePage") - 1
            End If
        Case ">"
            If Not rsPicktPick.EOF Then
                Session("rsPicktPick_AbsolutePage") = Session("rsPicktPick_AbsolutePage") + 1
            End If
        Case ">>"
            If fSupportsBookmarks Then
                Session("rsPicktPick_AbsolutePage") = rsPicktPick.PageCount
            End If
    End Select
    Do
        If fSupportsBookmarks Then
            rsPicktPick.AbsolutePage = Session("rsPicktPick_AbsolutePage")
        Else
            If fNeedRecordset Or fMoveAbsolute Or rsPicktPick.EOF Or Not fSupportsMovePrevious Then
                rsPicktPick.MoveFirst
                rsPicktPick.Move (Session("rsPicktPick_AbsolutePage") - 1) * tPageSize
            Else
                intCurPos = ((tPrevAbsolutePage - 1) * tPageSize) + tPageSize
                intNewPos = ((Session("rsPicktPick_AbsolutePage") - 1) * tPageSize) + 1
                rsPicktPick.Move intNewPos - intCurPos
            End If
            If rsPicktPick.BOF Then rsPicktPick.MoveNext
        End If
        If Not rsPicktPick.EOF Then Exit Do
        Session("rsPicktPick_AbsolutePage") = Session("rsPicktPick_AbsolutePage") - 1
    Loop
End If

Do
    If fEmptyRecordset Then Exit Do
    If tRecordsProcessed = tPageSize Then Exit Do
    If Not fFirstPass Then
        rsPicktPick.MoveNext
    Else
        fFirstPass = False
    End If
    If rsPicktPick.EOF Then Exit Do
    tRecordsProcessed = tRecordsProcessed + 1
%>
<!--METADATA TYPE="DesignerControl" endspan-->

<% 
If strFormMode = "Edit" Then
	Response.Write "<P>"
	Response.Write "<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=2 BORDER=0>"
	ShowField "PickID", "PickID", True, Null
	ShowField "MemberID", "MemberID", False, avarPicktMember
	ShowField "Pick", "Pick", False, Null
	ShowField "Comment", "Comment", False, Null
	ShowField "Date", "Date", False, Null
	Response.Write "</TABLE>"
	Response.Write "</FORM></P>"
	stQueryString = "?FormMode=Edit"
	fHideNavBar = False
	fHideRule = True
Else
	fHideNavBar = True
	fHideRule = True
End If 
%>

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

<% 
If strFormMode <> "Edit" Then
	Response.Write "<P>"
	Response.Write "<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=2 BORDER=0>"
	ShowField "PickID", "PickID", True, Null
	ShowField "MemberID", "MemberID", False, avarPicktMember
	ShowField "Pick", "Pick", False, Null
	ShowField "Comment", "Comment", False, Null
	ShowField "Date", "Date", False, Null
	Response.Write "</TABLE>"
	Response.Write "</FORM></P>"	
End If
%>

<!---------------------------- Footer Section --------------------------------->

<%
' Display a message if there are no records to show
If strFormMode = "Edit" And fEmptyRecordset Then
	Response.Write "<p align=left>No Records Available</p>"
End If
' TEMP: This is here until we get a drop of the data range that has
' 		the CacheRecordset property  
If fNeedRecordset Then
	Set Session("rsPicktPick_Recordset") = rsPicktPick
End If
%>

</body>
</html>

