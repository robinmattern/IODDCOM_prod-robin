<%@ LANGUAGE="VBScript" %>
<%
'-------------------------------------------------------------------------------
' Microsoft Visual InterDev - Data Form Wizard
' 
' Action Page
'
' (c) 1997 Microsoft Corporation.  All Rights Reserved.
'
' This file is an Active Server Page that contains the server script that 
' handles filter, update, insert, and delete commands from the form view of a 
' Data Form. It can also echo back confirmation of database operations and 
' report errors. Some commands are passed through and redirected. Microsoft 
' Internet Information Server 3.0 is required.
'
'-------------------------------------------------------------------------------

Dim strDFName
Dim strErrorAdditionalInfo
strDFName = "rsPantPan"
%>

<script RUNAT="Server" LANGUAGE="VBScript">

'---- FieldAttributeEnum Values ----
Const adFldUpdatable = &H00000004
Const adFldUnknownUpdatable = &H00000008
Const adFldIsNullable = &H00000020

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- DataTypeEnum Values ----
Const adUnsignedTinyInt = 17
Const adBoolean = 11
Const adDate = 7
Const adDBDate = 133
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205

'---- Error Values ----
Const errInvalidPrefix = 20001		'Invalid wildcard prefix
Const errInvalidOperator = 20002	'Invalid filtering operator
Const errInvalidOperatorUse = 20003	'Invalid use of LIKE operator
Const errNotEditable = 20011		'Field not editable
Const errValueRequired = 20012		'Value required

'-------------------------------------------------------------------------------
' Purpose:  Substitutes Null for Empty
' Inputs:   varTemp	- the target value
' Returns:	The processed value
'-------------------------------------------------------------------------------

Function RestoreNull(varTemp)
	If Trim(varTemp) = "" Then
		RestoreNull = Null
	Else
		RestoreNull = varTemp
	End If
End Function

Sub RaiseError(intErrorValue, strFieldName)
	Dim strMsg	
	Select Case intErrorValue
		Case errInvalidPrefix
			strMsg = "Wildcard characters * and % can only be used at the end of the criteria"
		Case errInvalidOperator
			strMsg = "Invalid filtering operators - use <= or >= instead."
		Case errInvalidOperatorUse
			strMsg = "The 'Like' operator can only be used with strings."
		Case errNotEditable
			strMsg = strFieldName & " field is not editable."
		Case errValueRequired
			strMsg = "A value is required for " & strFieldName & "."
	End Select
	Err.Raise intErrorValue, "DataForm", strMsg
End Sub

'-------------------------------------------------------------------------------
' Purpose:  Converts to subtype of string - handles Null cases
' Inputs:   varTemp	- the target value
' Returns:	The processed value
'-------------------------------------------------------------------------------

Function ConvertToString(varTemp)
	If IsNull(varTemp) Then
		ConvertToString = Null
	Else
		ConvertToString = CStr(varTemp)
	End If
End Function

'-------------------------------------------------------------------------------
' Purpose:  Tests to equality while dealing with Null values
' Inputs:   varTemp1	- the first value
'			varTemp2	- the second value
' Returns:	True if equal, False if not
'-------------------------------------------------------------------------------

Function IsEqual(ByVal varTemp1, ByVal varTemp2)
	IsEqual = False
	If IsNull(varTemp1) And IsNull(varTemp2) Then
		IsEqual = True
	Else
		If IsNull(varTemp1) Then Exit Function
		If IsNull(varTemp2) Then Exit Function
	End If
	If varTemp1 = varTemp2 Then IsEqual = True
End Function

'-------------------------------------------------------------------------------
' Purpose:  Tests whether the field in the recordset is required
' Assumes: 	That the recordset containing the field is open
' Inputs:   strFieldName	- the name of the field in the recordset
' Returns:	True if updatable, False if not
'-------------------------------------------------------------------------------

Function IsRequiredField(strFieldName)
	IsRequiredField = False
	If (rsPantPan(strFieldName).Attributes And adFldIsNullable) = 0 Then 
		IsRequiredField = True
	End If
End Function

'-------------------------------------------------------------------------------
' Purpose:  Tests whether the field in the recordset is updatable
' Assumes: 	That the recordset containing the field is open
' Effects:	Sets Err object if field is not updatable
' Inputs:   strFieldName	- the name of the field in the recordset
' Returns:	True if updatable, False if not
'-------------------------------------------------------------------------------

Function CanUpdateField(strFieldName)
	Dim intUpdatable
	intUpdatable = (adFldUpdatable Or adFldUnknownUpdatable)
	CanUpdateField = True
	If (rsPantPan(strFieldName).Attributes And intUpdatable) = False Then
		CanUpdateField = False
	End If
End Function

'-------------------------------------------------------------------------------
' Purpose:  Insert operation - updates a recordset field with a new value 
'			during an insert operation.
' Assumes: 	That the recordset containing the field is open
' Effects:	Sets Err object if field is not set but is required
' Inputs:   strFieldName	- the name of the field in the recordset
' Returns:	True if successful, False if not
'-------------------------------------------------------------------------------

Function InsertField(strFieldName)
	InsertField = True
	If IsEmpty(Request(strFieldName)) Then Exit Function
	Select Case rsPantPan(strFieldName).Type
 		Case adBinary, adVarBinary, adLongVarBinary		'Binary
		Case Else
			If CanUpdateField(strFieldName) Then
				If IsRequiredField(strFieldName) And IsNull(RestoreNull(Request(strFieldName))) Then
					RaiseError errValueRequired, strFieldName
					InsertField = False
					Exit Function
				End If				
				rsPantPan(strFieldName) = RestoreNull(Request(strFieldName))
			End If
	End Select
End Function

'-------------------------------------------------------------------------------
' Purpose:  Update operation - updates a recordset field with a new value 
' Assumes: 	That the recordset containing the field is open
' Effects:	Sets Err object if field is not set but is required
' Inputs:   strFieldName	- the name of the field in the recordset
' Returns:	True if successful, False if not
'-------------------------------------------------------------------------------

Function UpdateField(strFieldName)
	UpdateField = True
	If IsEmpty(Request(strFieldName)) Then Exit Function
	Select Case rsPantPan(strFieldName).Type
 		Case adBinary, adVarBinary, adLongVarBinary		'Binary
		Case Else
			' Only update if the value has changed
			If Not IsEqual(ConvertToString(rsPantPan(strFieldName)), RestoreNull(Request(strFieldName))) Then
				If CanUpdateField(strFieldName) Then						
					If IsRequiredField(strFieldName) And IsNull(RestoreNull(Request(strFieldName))) Then
						RaiseError errValueRequired, strFieldName
						UpdateField = False
						Exit Function
					End If				
					rsPantPan(strFieldName) = RestoreNull(Request(strFieldName))
				Else
					RaiseError errNotEditable, strFieldName
					UpdateField = False
				End If
			End If
	End Select
End Function

'-------------------------------------------------------------------------------
' Purpose:  Criteria handler for a field in the recordset. Determines
'			correct delimiter based on data type
' Effects:	Appends to strWhere and strWhereDisplay variables
' Inputs:   strFieldName	- the name of the field in the recordset
'			avarLookup		- lookup array - null if none
'-------------------------------------------------------------------------------

Sub FilterField(ByVal strFieldName, avarLookup)
	Dim strFieldDelimiter
	Dim strDisplayValue
	Dim strValue
	Dim intRow
	strValue = Request(strFieldName)
	strDisplayValue = Request(strFieldName)
	
	' If empty then exit right away
	If Request(strFieldName) = "" Then Exit Sub
	
	' Concatenate the And boolean operator
	If strWhere <> "" Then strWhere = strWhere & " And"
	If strWhereDisplay <> "" Then strWhereDisplay = strWhereDisplay & " And"
	
	' If lookup field, then use lookup value for display
	If Not IsNull(avarLookup) Then
		For intRow = 0 to UBound(avarLookup, 2)
			If CStr(avarLookup(0, intRow)) = Request(strFieldName) Then
				strDisplayValue = avarLookup(1, intRow)
				Exit For
			End If
		Next
	End If
	
	' Set delimiter based on data type
	Select Case rsPantPan(strFieldName).Type
		Case adBSTR, adChar, adWChar, adVarChar, adVarWChar	'string types
			strFieldDelimiter = "'"
		Case adLongVarChar, adLongVarWChar					'long string types
			strFieldDelimiter = "'"				
		Case adDate, adDBDate, adDBTimeStamp				'date types
			strFieldDelimiter = "#"
		Case Else
			strFieldDelimiter = ""
	End Select
	
	' Modifies script level variables
	strWhere = strWhere & " " & PrepFilterItem(strFieldName, strValue, strFieldDelimiter)
	strWhereDisplay = strWhereDisplay & " " & PrepFilterItem(strFieldName, strDisplayValue, strFieldDelimiter)

End Sub

'-------------------------------------------------------------------------------
' Purpose:  Constructs a name/value pair for a where clause
' Effects:	Sets Err object if the criteria is invalid
' Inputs:   strFieldName	- the name of the field in the recordset
'			strCriteria		- the criteria to use
'			strDelimiter	- the proper delimiter to use
' Returns:	The name/value pair as a string
'-------------------------------------------------------------------------------

Function PrepFilterItem(ByVal strFieldName, ByVal strCriteria, ByVal strDelimiter)
	Dim strOperator
	Dim intEndOfWord
	Dim strWord

	' Char, VarChar, and LongVarChar must be single quote delimited.
	' Dates are pound sign delimited.
	' Numerics should not be delimited.
	' String to Date conversion rules are same as VBA.
	' Only support for ANDing.
	' Support the LIKE operator but only with * or % as suffix.
	
	strCriteria = Trim(strCriteria)	'remove leading/trailing spaces
	strOperator = "="				'sets default
	strValue = strCriteria			'sets default

	' Get first word and look for operator
	intEndOfWord = InStr(strCriteria, " ")
	If intEndOfWord Then
		strWord = UCase(Left(strCriteria, intEndOfWord - 1))
		' See if the word is an operator
		Select Case strWord
			Case "=", "<", ">", "<=", ">=",  "<>", "LIKE"
				strOperator = strWord
				strValue = Trim(Mid(strCriteria, intEndOfWord + 1))
			Case "=<", "=>"
				RaiseError errInvalidOperator, strFieldName
		End Select
	Else
		strWord = UCase(Left(strCriteria, 2))
		Select Case strWord
			Case "<=", ">=", "<>"
				strOperator = strWord
				strValue = Trim(Mid(strCriteria, 3))
			Case "=<", "=>"
				RaiseError errInvalidOperator, strFieldName
			Case Else
				strWord = UCase(Left(strCriteria, 1))
				Select Case strWord
					Case "=", "<", ">"
						strOperator = strWord
						strValue = Trim(Mid(strCriteria, 2))
				End Select
		End Select
	End If

	' Make sure LIKE is only used with strings
	If strOperator = "LIKE" and strDelimiter <> "'" Then
		RaiseError errInvalidOperatorUse, strFieldName
	End If		

	' Strip any extraneous delimiters because we add them anyway
	' Single Quote
	If Left(strValue, 1) = Chr(39) Then strValue = Mid(strValue, 2)
	If Right(strValue, 1) = Chr(39) Then strValue = Left(strValue, Len(strValue) - 1)

	' Double Quote - just in case
	If Left(strValue, 1) = Chr(34) Then strValue = Mid(strValue, 2)
	If Right(strValue, 1) = Chr(34) Then strValue = Left(strValue, Len(strValue) - 1)

	' Pound sign - dates
	If Left(strValue, 1) = Chr(35) Then strValue = Mid(strValue, 2)
	If Right(strValue, 1) = Chr(35) Then strValue = Left(strValue, Len(strValue) - 1)
    
	' Check for leading wildcards
	If Left(strValue, 1) = "*" Or Left(strValue, 1) = "%" Then
		RaiseError errInvalidPrefix, strFieldName
	End If
	
	PrepFilterItem = "[" & strFieldName & "]" & " " & strOperator & " " & strDelimiter & strValue & strDelimiter
End Function

'-------------------------------------------------------------------------------
' Purpose:  Display field involved in a database operation for feedback.
' Assumes: 	That the recordset containing the field is open
' Inputs:   strFieldLabel	- the label to be used for the field
'			strFieldName	- the name of the field in the recordset
'-------------------------------------------------------------------------------

Sub FeedbackField(strFieldLabel, strFieldName, avarLookup)
	Dim strBool
	Dim intRow
	Response.Write "<TR VALIGN=TOP>"
    Response.Write "<TD ALIGN=Left><FONT SIZE=-1><B>&nbsp;&nbsp;" & strFieldLabel & "</B></FONT></TD>"
	Response.Write "<TD BGCOLOR=White WIDTH=100% ALIGN=Left><FONT SIZE=-1>"
	
	' Test for lookup
	If Not IsNull(avarLookup) Then
		For intRow = 0 to UBound(avarLookup, 2)
			If CStr(avarLookup(0, intRow)) = Request(strFieldName) Then
				Response.Write Server.HTMLEncode(avarLookup(1, intRow))
				Exit For
			End If
		Next
		Response.Write "</FONT></TD></TR>"
		Exit Sub
	End If
	
	' Test for empty
	If Request(strFieldName) = "" Then
		Response.Write "&nbsp;"
		Response.Write "</FONT></TD></TR>"
		Exit Sub
	End If
	
	' Test the data types and display appropriately	
	Select Case rsPantPan(strFieldName).Type
		Case adBoolean, adUnsignedTinyInt				'Boolean
			strBool = ""
			If Request(strFieldName) <> 0 Then
				strBool = "True"
			Else
				strBool = "False"
			End If
			Response.Write strBool
		Case adBinary, adVarBinary, adLongVarBinary		'Binary
			Response.Write "[Binary]"
		Case adLongVarChar, adLongVarWChar				'Memo
			Response.Write Server.HTMLEncode(Request(strFieldName))
		Case Else
			If Not CanUpdateField(strFieldName) Then
				Response.Write "[AutoNumber]"
			Else
				Response.Write Server.HTMLEncode(Request(strFieldName))
			End If
	End Select
	Response.Write "</FONT></TD></TR>"
End Sub

</script>

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


<% 
If Not IsEmpty(Request("DataAction")) Then
	strDataAction = Trim(Request("DataAction"))
Else
	Response.Redirect "PanForm.asp?FormMode=Edit"
End If

'------------------
' Action handler
'------------------
Select Case strDataAction
	
	Case "List View"
		
		Response.Redirect "PanList.asp"

	Case "Cancel"

		Response.Redirect "PanForm.asp?FormMode=Edit"

	Case "Filter"
	
		On Error Resume Next
		Session("rsPantPan_Filter") = ""
		Session("rsPantPan_FilterDisplay") = ""
		Session("rsPantPan_Recordset").Filter = ""
		Response.Redirect "PanForm.asp?FormMode=" & strDataAction

	Case "New"
	
		On Error Resume Next
		Session("rsPantPan_Filter") = ""
		Session("rsPantPan_FilterDisplay") = ""
		Session("rsPantPan_Recordset").Filter = ""
		Response.Redirect "PanForm.asp?FormMode=" & strDataAction

	Case "Find"

		Session("rsPantPan_PageSize") = 1 'So we don't do standard page conversion
		Session("rsPantPan_AbsolutePage") = CLng(Request("Bookmark"))
		Response.Redirect "PanForm.asp"

	Case "All Records"
	
		On Error Resume Next
		Session("rsPantPan_Filter") = ""
		Session("rsPantPan_FilterDisplay") = ""
		Session("rsPantPan_Recordset").Filter = ""
		Session("rsPantPan_AbsolutePage") = 1
		Response.Redirect "PanForm.asp"

	Case "Apply"

		On Error Resume Next
		
		' Make sure we exit and re-process the form if session has timed out
		If IsEmpty(Session("rsPantPan_Recordset")) Then
			Response.Redirect "PanForm.asp?FormMode=Edit"
		End If
		
		Set rsPantPan = Session("rsPantPan_Recordset")

		strWhere = ""
		strWhereDisplay = ""
		FilterField "PanID", Null
		FilterField "MemberID", avarPantMember
		FilterField "Pan", Null
		FilterField "Comment", Null
		FilterField "Date", Null
        
		' Filter the recordset
		If strWhere <> "" Then
			Session("rsPantPan_Filter") = strWhere
			Session("rsPantPan_FilterDisplay") = strWhereDisplay
			Session("rsPantPan_AbsolutePage") = 1
		Else
			Session("rsPantPan_Filter") = ""
			Session("rsPantPan_FilterDisplay") = ""
		End If

		' Jump back to the form
		If Err.Number = 0 Then Response.Redirect "PanForm.asp"

	Case "Insert"

		On Error Resume Next		

		' Make sure we exit and re-process the form if session has timed out
		If IsEmpty(Session("rsPantPan_Recordset")) Then
			Response.Redirect "PanForm.asp?FormMode=Edit"
		End If
		
		Set rsPantPan = Session("rsPantPan_Recordset")
		rsPantPan.AddNew
		
		Do
			If Not InsertField("PanID") Then Exit Do
			If Not InsertField("MemberID") Then Exit Do
			If Not InsertField("Pan") Then Exit Do
			If Not InsertField("Comment") Then Exit Do
			If Not InsertField("Date") Then Exit Do

			rsPantPan.Update
			Exit Do
		Loop

		If Err.Number <> 0 Then
			If rsPantPan.EditMode Then rsPantPan.CancelUpdate
		Else
			If IsEmpty(Session("rsPantPan_AbsolutePage")) Or Session("rsPantPan_AbsolutePage") = 0 Then
				Session("rsPantPan_AbsolutePage") = 1
			End If
			' Requery static cursor so inserted record is visible
			If rsPantPan.CursorType = adOpenStatic Then rsPantPan.Requery
			Session("rsPantPan_Status") = "Record has been inserted"
		End If

	Case "Update"

		On Error Resume Next		

		' Make sure we exit and re-process the form if session has timed out
		If IsEmpty(Session("rsPantPan_Recordset")) Then
			Response.Redirect "PanForm.asp?FormMode=Edit"
		End If
		
		Set rsPantPan = Session("rsPantPan_Recordset")
		If rsPantPan.EOF and rsPantPan.BOF Then Response.Redirect "PanForm.asp"
		
		Do

			If Not UpdateField("PanID") Then Exit Do
			If Not UpdateField("MemberID") Then Exit Do
			If Not UpdateField("Pan") Then Exit Do
			If Not UpdateField("Comment") Then Exit Do
			If Not UpdateField("Date") Then Exit Do

			If rsPantPan.EditMode Then rsPantPan.Update
			Exit Do
		Loop

		If Err.Number <> 0 Then
			If rsPantPan.EditMode Then rsPantPan.CancelUpdate
		End If

	Case "Delete"

		On Error Resume Next
		
		' Make sure we exit and re-process the form if session has timed out
		If IsEmpty(Session("rsPantPan_Recordset")) Then
			Response.Redirect "PanForm.asp?FormMode=Edit"
		End If
		
		Set rsPantPan = Session("rsPantPan_Recordset")
		If rsPantPan.EOF and rsPantPan.BOF Then Response.Redirect "PanForm.asp"
		
		rsPantPan.Delete

		' Proceed if no error
		If Err.Number = 0 Then
			' Requery static cursor so deleted record is removed
			If rsPantPan.CursorType = adOpenStatic Then rsPantPan.Requery
			
			' Move off deleted rec
			rsPantPan.MoveNext
			
			' If at EOF then jump back one and adjust AbsolutePage
			If rsPantPan.EOF Then
				Session("rsPantPan_AbsolutePage") = Session("rsPantPan_AbsolutePage") - 1				
				If rsPantPan.BOF And rsPantPan.EOF Then rsPantPan.Requery
			End If
		End If

End Select
%>
<%
'<!----------------------------- Error Handler --------------------------------->

   If Err Then %>
	<%
	' Add additional error information to clarify specific errors
	Select Case Err.Number
		Case -2147467259
			strErrorAdditionalInfo = "  This may be caused by an attempt to update a non-primary table in a view."
		Case Else
			strErrorAdditionalInfo = ""
	End Select
	%>
	<html>
	<head>
		<meta NAME="GENERATOR" CONTENT="Microsoft Visual InterDev">
		
		<meta NAME="keywords" CONTENT="Microsoft Data Form, tPan Form">
		<title>tPan Form</title>
	<meta name="Microsoft Theme" content="none, default">
</head>
	<basefont FACE="Arial, Helvetica, sans-serif">
	<body>
	<table WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
		<tr>
			<th COLSPAN="2" NOWRAP ALIGN="Left" BGCOLOR="green">
				<font SIZE="5">&nbsp;Message:&nbsp;</font>
			</th>
		</tr>
		<tr>
			<td BGCOLOR="#FFFFCC" COLSPAN="2">
			<font SIZE="3"><b>
			<% 
			Select Case strDataAction
				Case "Insert"
					Response.Write("Unable to insert the record into tPan.")
				Case "Update"
					Response.Write("Unable to post the updated record to tPan.")
				Case "Delete"
					Response.Write("Unable to delete the record from tPan.")
			End Select
			%>
			</b></font>
			</td>
		</tr>
	</table>
	<table WIDTH="100%" CELLSPACING="1" CELLPADDING="2" BORDER="0">
		<tr>
			<td ALIGN="Left" BGCOLOR="green"><font SIZE="-1"><b>&nbsp;&nbsp;Item</b></font></td>
			<td WIDTH="100%" ALIGN="Left" BGCOLOR="green"><font SIZE="-1"><b>Description</b></font></td>
		</tr>
		<tr>
			<td><font SIZE="-1"><b>&nbsp;&nbsp;Source:</b></font></td>
			<td BGCOLOR="White"><font SIZE="-1"><%= Err.Source %></td>
		</tr>
		<tr>
			<td NOWRAP><font SIZE="-1"><b>&nbsp;&nbsp;Error Number:</b></font></td>
			<td BGCOLOR="White"><font SIZE="-1"><%= Err.Number %></font></td>
		</tr>
		<tr>
			<td><font SIZE="-1"><b>&nbsp;&nbsp;Description:</b></font></td>
			<td BGCOLOR="White"><font SIZE="-1"><%= Server.HTMLEncode(Err.Description & strErrorAdditionalInfo) %></font></td>
		</tr>
		<tr>
			<td COLSPAN="2"><hr></td>
		</tr>
		<tr>
			<td>
			<% Response.Write "<FORM ACTION=""PanForm.asp"" METHOD=""POST"">" %>
			<input TYPE="Hidden" NAME="FormMode" VALUE="Edit">
			<input TYPE="SUBMIT" VALUE="Form View">
			</form>
			</td>
			<td>
			<font SIZE="-1">
			To return to the form view with the previously entered 
			information intact, use your browsers &quot;back&quot; button
			</font>
			</td>
		</tr>
	</table>
	</body>
	</html>

<% Else %>
<!-- Action feedback -->
	<html>
	<head>
		<meta NAME="GENERATOR" CONTENT="Microsoft Visual InterDev">
		
		<meta NAME="keywords" CONTENT="Microsoft DataForm, tPan Form">
		<title>tPan Form</title>
	</head>
	<basefont FACE="Arial, Helvetica, sans-serif">
	<body>
	<table WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
		<tr>
			<th COLSPAN="2" NOWRAP ALIGN="Left" BGCOLOR="#CCCCFF">
				<font SIZE="5">&nbsp;Feedback:&nbsp;</font>
			</th>
		</tr>
		<tr>
			<td BGCOLOR="#FFFFCC" COLSPAN="2">&nbsp;&nbsp;
			<font SIZE="-1">
			<% 
			Select Case strDataAction
				Case "Insert"
					Response.Write("The following record has been inserted into PantPan.")
				Case "Update"
					Response.Write("The following updated record has been posted to PantPan.")
				Case "Delete"
					Response.Write("The following record has been deleted from PantPan.")
			End Select
			%>
			</font>
			</td>
		</tr>
	</table>
	<table WIDTH="100%" CELLSPACING="1" CELLPADDING="2" BORDER="0">
		<tr>
			<td WIDTH="30%" ALIGN="Left" BGCOLOR="green"><font SIZE="-1"><b>&nbsp;&nbsp;Field</b></font></td>
			<td WIDTH="100%" ALIGN="Left" BGCOLOR="green"><font SIZE="-1"><b>Value</b></font></td>
		</tr>
		<%
			FeedbackField "PanID", "PanID", Null
			FeedbackField "MemberID", "MemberID", avarPantMember
			FeedbackField "Pan", "Pan", Null
			FeedbackField "Comment", "Comment", Null
			FeedbackField "Date", "Date", Null
		%>
		<tr>
			<td COLSPAN="2"><hr></td>
		</tr>
		<tr>
			<td COLSPAN="2">
			<% Response.Write "<FORM ACTION=""PanForm.asp"" METHOD=""POST"">" %>
				<% If strDataAction = "Insert" Then %>
					<input TYPE="SUBMIT" NAME="FormMode" VALUE="New">
					<input TYPE="SUBMIT" NAME="FormMode" VALUE="Form View">
				<% Else %>
					<input TYPE="Hidden" NAME="FormMode" VALUE="Edit">
					<input TYPE="SUBMIT" VALUE="Form View">
				<% End If %>
			</form>
			</td>
		</tr>
	</table>
</body>
</html>

<% 
End If 
Set rsPantPan = Nothing
%>

