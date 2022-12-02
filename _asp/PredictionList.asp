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
strDFName = "rsPredictiontPrediction"
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
			If ConvertNull(avarLookup(0, intRow)) = ConvertNull(rsPredictiontPrediction(strFieldName)) Then
				Response.Write Server.HTMLEncode(ConvertNull(avarLookup(1, intRow)))
				Exit For
			End If
		Next
		Response.Write "</FONT></TD>"
		Exit Function
	End If
	
	Select Case rsPredictiontPrediction(strFieldName).Type
		Case adBoolean, adUnsignedTinyInt				'Boolean
			strBool = ""
			If rsPredictiontPrediction(strFieldName) <> 0 Then
				strBool = "True"
			Else
				strBool = "False"
			End If
			Response.Write "<TD BGCOLOR=White ><FONT SIZE=-1>" & strBool & "</FONT></TD>"
			
		Case adBinary, adVarBinary, adLongVarBinary		'Binary
			Response.Write "<TD BGCOLOR=White ><FONT SIZE=-1>[Binary]</FONT></TD>"
			
		Case adLongVarChar, adLongVarWChar				'Memo
			Response.Write "<TD BGCOLOR=White NOWRAP><FONT SIZE=-1>"
			strPartial = Left(rsPredictiontPrediction(strFieldName), 50)			
			If ConvertNull(strPartial) = "" Then
				Response.Write "&nbsp;"
			Else
				Response.Write Server.HTMLEncode(strPartial)
			End If
			If rsPredictiontPrediction(strFieldName).ActualSize > 50 Then Response.Write "..."
			Response.Write "</FONT></TD>"
			
		Case Else
			Response.Write "<TD BGCOLOR=White ALIGN=Left NOWRAP><FONT SIZE=-1>"
			If ConvertNull(rsPredictiontPrediction(strFieldName)) = "" Then
				Response.Write "&nbsp;"
			Else
				' Check for special field types
				Select Case UCase(Left(rsPredictiontPrediction(strFieldName).Name, 4))
					Case "URL_"
						Response.Write "<A HREF=" & QuotedString(rsPredictiontPrediction(strFieldName)) & ">"
						Response.Write Server.HTMLEncode(ConvertNull(rsPredictiontPrediction(strFieldName)))
						Response.Write "</A>"
					Case Else
						If IsURL(rsPredictiontPrediction(strFieldName)) Then
							Response.Write "<A HREF=" & QuotedString(rsPredictiontPrediction(strFieldName)) & ">"
							Response.Write Server.HTMLEncode(ConvertNull(rsPredictiontPrediction(strFieldName)))
							Response.Write "</A>"
						Else
							Response.Write Server.HTMLEncode(ConvertNull(rsPredictiontPrediction(strFieldName)))
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
	
	<meta NAME="Keywords" CONTENT="Microsoft Data Form, tPrediction List">
	<title>tPrediction List</title>
<% ' FP_ASP -- ASP Automatically generated by a Frontpage Component. Do not Edit.
FP_CharSet = "windows-1252"
FP_CodePage = 1252 %>
</head>

<!--------------------------- Formatting Section ------------------------------>

<basefont FACE="Arial, Helvetica, sans-serif">
<body>

<!---------------------------- Lookups Section ------------------------------->
<% 
Dim avarPredictiontMember
If IsEmpty(Application(strDFName & "_Lookup_PredictiontMember")) Or strPagingMove = "Requery" Then
    Set DataConn = Server.CreateObject("ADODB.Connection")
    Dataconn.ConnectionTimeout = Session("DataConn_ConnectionTimeout")
    Dataconn.CommandTimeout = Session("DataConn_CommandTimeout")
    Dataconn.Open Session("DataConn_ConnectionString"), Session("DataConn_RuntimeUserName"), Session("DataConn_RuntimePassword")
	Set rsPredictiontMember = Dataconn.Execute("SELECT DISTINCT `MemberID`, `LastName` FROM `tMember`")
	avarPredictiontMember = Null
	On Error Resume Next
	avarPredictiontMember = rsPredictiontMember.GetRows()
	Application.Lock
	Application(strDFName & "_Lookup_PredictiontMember") = avarPredictiontMember
	Application.Unlock
Else
	avarPredictiontMember = Application(strDFName & "_Lookup_PredictiontMember")
End If
%>


<!---------------------------- Heading Section ------------------------------->

<% Response.Write "<FORM ACTION=PredictionForm.asp METHOD=""POST"">" %>
<table WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
	<tr>
		<th NOWRAP BGCOLOR="Green" ALIGN="Left">
			<font SIZE="5">&nbsp;Predictions</font>
		</th>
		<td BGCOLOR="Green" VALIGN="Middle" ALIGN="Right" WIDTH="100%">
			<input TYPE="Hidden" NAME="FormMode" VALUE="Edit">
			&nbsp;<input TYPE="SUBMIT" NAME="DataAction" VALUE="Form View">&nbsp;
		</td>
	</tr>
	<tr>
		<td BGCOLOR="#FFFFCC" COLSPAN="3">
			<font SIZE="-1">&nbsp;&nbsp;
			<% 
			If IsEmpty(Session("rsPredictiontPrediction_Filter")) Or Session("rsPredictiontPrediction_Filter")="" Then
				Response.Write "Current Filter: None<BR>"
			Else
				Response.Write "Current Filter: " & Session("rsPredictiontPrediction_FilterDisplay") & "<BR>"
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
	<tr BGCOLOR="Green" VALIGN="TOP">
		<td ALIGN="Center"><font SIZE="-1">&nbsp;#&nbsp;</font></td>
		<td ALIGN="Left"><font SIZE="-1"><b>PredictionID</b></font></td>
		<td ALIGN="Left"><font SIZE="-1"><b>MemberID</b></font></td>
		<td ALIGN="Left"><font SIZE="-1"><b>Prediction</b></font></td>
		<td ALIGN="Left"><font SIZE="-1"><b>Date</b></font></td>
	</tr>

<!--METADATA TYPE="DesignerControl" startspan
	<OBJECT ID="rsPredictiontPrediction" WIDTH=151 HEIGHT=24
		CLASSID="CLSID:F602E721-A281-11CF-A5B7-0080C73AAC7E">
		<PARAM NAME="BarAlignment" VALUE="0">
       	<PARAM NAME="PageSize" VALUE="15">
		<PARAM Name="RangeType" Value="2">
		<PARAM Name="DataConnection" Value="DataConn">
		<PARAM Name="CommandType" Value="0">
		<PARAM Name="CommandText" Value="SELECT `PredictionID`, `MemberID`, `Prediction`, `Date` FROM `tPrediction`">
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
tHeaderName = "rsPredictiontPrediction"
tPageSize = 15
tPagingMove = ""
tRangeType = "Table"
tRecordsProcessed = 0
tPrevAbsolutePage = 0
intCurPos = 0
intNewPos = 0
fSupportsBookmarks = True
fMoveAbsolute = False

If Not IsEmpty(Request("rsPredictiontPrediction_PagingMove")) Then
    tPagingMove = Trim(Request("rsPredictiontPrediction_PagingMove"))
End If

If IsEmpty(Session("rsPredictiontPrediction_Recordset")) Then
    fNeedRecordset = True
Else
    If Session("rsPredictiontPrediction_Recordset") Is Nothing Then
        fNeedRecordset = True
    Else
        Set rsPredictiontPrediction = Session("rsPredictiontPrediction_Recordset")
    End If
End If

If fNeedRecordset Then
    Set DataConn = Server.CreateObject("ADODB.Connection")
    Dataconn.ConnectionTimeout = Session("DataConn_ConnectionTimeout")
    Dataconn.CommandTimeout = Session("DataConn_CommandTimeout")
    Dataconn.Open Session("DataConn_ConnectionString"), Session("DataConn_RuntimeUserName"), Session("DataConn_RuntimePassword")
    Set cmdTemp = Server.CreateObject("ADODB.Command")
    Set rsPredictiontPrediction = Server.CreateObject("ADODB.Recordset")
    cmdTemp.CommandText = "SELECT `PredictionID`, `MemberID`, `Prediction`, `Date` FROM `tPrediction`"
    cmdTemp.CommandType = 1
    Set cmdTemp.ActiveConnection = DataConn
    rsPredictiontPrediction.Open cmdTemp, , 1, 3
End If
On Error Resume Next
If rsPredictiontPrediction.BOF And rsPredictiontPrediction.EOF Then fEmptyRecordset = True
On Error Goto 0
If Err Then fEmptyRecordset = True
If fNeedRecordset Then
    Set Session("rsPredictiontPrediction_Recordset") = rsPredictiontPrediction
End If
rsPredictiontPrediction.PageSize = tPageSize
fSupportsBookmarks = rsPredictiontPrediction.Supports(8192)

If Not IsEmpty(Session("rsPredictiontPrediction_Filter")) And Not fEmptyRecordset Then
    rsPredictiontPrediction.Filter = Session("rsPredictiontPrediction_Filter")
    If rsPredictiontPrediction.BOF And rsPredictiontPrediction.EOF Then fEmptyRecordset = True
End If

If IsEmpty(Session("rsPredictiontPrediction_PageSize")) Then Session("rsPredictiontPrediction_PageSize") = tPageSize
If IsEmpty(Session("rsPredictiontPrediction_AbsolutePage")) Then Session("rsPredictiontPrediction_AbsolutePage") = 1

If Session("rsPredictiontPrediction_PageSize") <> tPageSize Then
    tCurRec = ((Session("rsPredictiontPrediction_AbsolutePage") - 1) * Session("rsPredictiontPrediction_PageSize")) + 1
    tNewPage = Int(tCurRec / tPageSize)
    If tCurRec Mod tPageSize <> 0 Then
        tNewPage = tNewPage + 1
    End If
    If tNewPage = 0 Then tNewPage = 1
    Session("rsPredictiontPrediction_PageSize") = tPageSize
    Session("rsPredictiontPrediction_AbsolutePage") = tNewPage
End If

If fEmptyRecordset Then
    fHideNavBar = True
    fHideRule = True
Else
    tPrevAbsolutePage = Session("rsPredictiontPrediction_AbsolutePage")
    Select Case tPagingMove
        Case ""
            fMoveAbsolute = True
        Case "Requery"
            rsPredictiontPrediction.Requery
            fMoveAbsolute = True
        Case "<<"
            Session("rsPredictiontPrediction_AbsolutePage") = 1
        Case "<"
            If Session("rsPredictiontPrediction_AbsolutePage") > 1 Then
                Session("rsPredictiontPrediction_AbsolutePage") = Session("rsPredictiontPrediction_AbsolutePage") - 1
            End If
        Case ">"
            If Not rsPredictiontPrediction.EOF Then
                Session("rsPredictiontPrediction_AbsolutePage") = Session("rsPredictiontPrediction_AbsolutePage") + 1
            End If
        Case ">>"
            If fSupportsBookmarks Then
                Session("rsPredictiontPrediction_AbsolutePage") = rsPredictiontPrediction.PageCount
            End If
    End Select
    Do
        If fSupportsBookmarks Then
            rsPredictiontPrediction.AbsolutePage = Session("rsPredictiontPrediction_AbsolutePage")
        Else
            If fNeedRecordset Or fMoveAbsolute Or rsPredictiontPrediction.EOF Or Not fSupportsMovePrevious Then
                rsPredictiontPrediction.MoveFirst
                rsPredictiontPrediction.Move (Session("rsPredictiontPrediction_AbsolutePage") - 1) * tPageSize
            Else
                intCurPos = ((tPrevAbsolutePage - 1) * tPageSize) + tPageSize
                intNewPos = ((Session("rsPredictiontPrediction_AbsolutePage") - 1) * tPageSize) + 1
                rsPredictiontPrediction.Move intNewPos - intCurPos
            End If
            If rsPredictiontPrediction.BOF Then rsPredictiontPrediction.MoveNext
        End If
        If Not rsPredictiontPrediction.EOF Then Exit Do
        Session("rsPredictiontPrediction_AbsolutePage") = Session("rsPredictiontPrediction_AbsolutePage") - 1
    Loop
End If

Do
    If fEmptyRecordset Then Exit Do
    If tRecordsProcessed = tPageSize Then Exit Do
    If Not fFirstPass Then
        rsPredictiontPrediction.MoveNext
    Else
        fFirstPass = False
    End If
    If rsPredictiontPrediction.EOF Then Exit Do
    tRecordsProcessed = tRecordsProcessed + 1
%>
<!--METADATA TYPE="DesignerControl" endspan-->

	<tr VALIGN="TOP">
		<td BGCOLOR="White"><font SIZE="-1">
		<%
		If tPageSize > 0 Then
			tCurRec = ((Session("rsPredictiontPrediction_AbsolutePage") - 1) * tPageSize) + tRecordsProcessed
		Else
			tRecordsProcessed = tRecordsProcessed + 1
			tCurRec = tRecordsProcessed
		End If
		Response.Write "<A HREF=" & QuotedString("PredictionAction.asp?Bookmark=" & tCurRec & "&DataAction=Find") & ">" & tCurRec & "</A>"
		%>

		</font></td>
		<%
		ShowField "PredictionID", Null
		ShowField "MemberID",avarPredictiontMember
		ShowField "Prediction", Null
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
	Set Session("rsPredictiontPrediction_Recordset") = rsPredictiontPrediction
End If 
%>

</td></tr></table>
</body>
</html>





