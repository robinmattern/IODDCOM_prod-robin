<%
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
Const adWChar = 130
Const adVarChar = 200
Const adLongVarChar = 201
Const adWVarChar = 202
Const adLongVarWChar = 203

Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205

'---- Other Values ----
Const dfMaxSize = 100

'---- Error Values ----
Const errInvalidPrefix = 20001		'Invalid wildcard prefix
Const errInvalidOperator = 20002	'Invalid filtering operator
Const errInvalidOperatorUse = 20003	'Invalid use of LIKE operator
Const errNotEditable = 20011		'Field not editable
Const errValueRequired = 20012		'Value required

%>
<%

Sub ClearVerifyDates()
Set CMS = Server.CreateObject("ADODB.Connection")
CMS.ConnectionTimeout = Session("CMS_ConnectionTimeout")
CMS.CommandTimeout = Session("CMS_CommandTimeout")
CMS.Open Session("CMS_ConnectionString"), Session("CMS_RuntimeUserName"), Session("CMS_RuntimePassword")
sqltext = "UPDATE tblCommittees SET DFODATE = NULL, GFODATE = NULL, CMODATE = NULL, DFOSignOff = 0,GFOSignOff = 0,CMOSignOff = 0, LastUpdated = GETDATE() WHERE "& SESSION("CIDstrWhere")

'CMS.BeginTrans
Set rsApprovals = CMS.Execute(sqltext)
'If CMS.Errors.Count = 0 then
'	CMS.CommitTrans
'Else
'	CMS.Rollback
'End If
'SET CMS = Nothing
Set rsApprovals = Nothing
SESSION("CommCanUpdate")=0
SESSION("CMOSignOff")=0

End Sub
%>

<%
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
%>

<%
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
%>

<%
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
%>

<%
'-------------------------------------------------------------------------------
' Purpose:  Tests whether the field in the recordset is required
' Assumes: 	That the recordset containing the field is open
' Inputs:   strFieldName	- the name of the field in the recordset
' Returns:	True if updatable, False if not
'-------------------------------------------------------------------------------

Function IsRequiredField(strFieldName)
	IsRequiredField = False
	If (rsAgenciestblAgencies(strFieldName).Attributes And adFldIsNullable) = 0 Then 
		IsRequiredField = True
	End If
End Function

%>

<%
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
%>

<%
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
%>

<%
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
%>

<%
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
%>

