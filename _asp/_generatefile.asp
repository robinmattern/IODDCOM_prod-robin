<%@ LANGUAGE="VBScript" %>
<!--#INCLUDE FILE="_incsessioncheck.asp"-->
<%'----- ASP Buffer Ouput, the server does not send a response to the client until all of the server scripts on the current page have been processed%>
<% 
Response.Buffer = True 
%>
<% Server.ScriptTimeOut = 6000 %>

<%'----- Expire the Page and Check that the User is currently Logged In. %>
<%
'<!--#INCLUDE FILE="_incexpires.asp"-->
'<!--#INCLUDE FILE="_incsessioncheck.asp"-->
%>
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<%
If len(trim(Session("SQLStr"))) = 0 then
	Session("Message") = "Report error: Sqlstr is empty."
	response.redirect "_message.asp"
End if
'response.write Session("SQLStr")

Set rs = Conn.Execute(Session("SQLStr"))
If rs.EOF then
	Session("Message") = "Report message: No records found."
	response.redirect "_message.asp"
End if	

SepType = Ucase(Trim(request.queryString("Type")))
If IsEmpty(SepType) OR Instr("CSV,TAB,COMMA,SPACE,FIXED,LABEL,XML",Ucase(SepType))=0 then
	SepType = "CSV"
End If
Response.Clear  ' Clear out buffer
Select Case Trim(Ucase(SepType))
Case "CSV"	
	response.contentType = "text/comma-separated-values"
	response.addHeader "Content-Disposition", "inline; filename=REPORT_CSV.csv;"
	Separator = "," 'Comma
Case "XML"	
	If Session("OKforXML") = "No" then
		Session("Message") = "Your browser, " & trim(Session("Browser")) & ", Does not display XML. Please use IE 5 or higher."
		response.redirect "_message.asp"
	End if
	response.contentType = "text/xml"
'	response.addHeader "Content-Disposition", "filename='REPORT.xml';"
	Separator = "" 'Not required
Case "TAB"	
	response.contentType = "text/comma-separated-values"
	response.addHeader "Content-Disposition", "inline; filename=REPORT_TAB.wri;"
	Separator = Chr(9) 'Tab
Case "COMMA"	
	response.contentType = "text/comma-separated-values"
	response.addHeader "Content-Disposition", "inline; filename=REPORT_COMMA.wri;"
	Separator = "," 'Comma
Case "SPACE"	
	response.contentType = "text/comma-separated-values"
	response.addHeader "Content-Disposition", "inline; filename=REPORT_SPACE.wri;"
	Separator = " " 'Space
Case "FIXED"	
	response.contentType = "text/comma-separated-values"
	response.addHeader "Content-Disposition", "inline; filename=REPORT_FIXED.wri;"
Case "LABEL"	
	response.contentType = "text/comma-separated-values"
	response.addHeader "Content-Disposition", "inline; filename=REPORT_LABELS.txt;"
End Select

Select Case Trim(Ucase(SepType))
Case "XML" 
	Session ("XMLFileString") = ""
	Session ("XMLFileString")=XMLfromSQL("QUERY", session("ReportName"),UniqueID())
	response.write Session("XMLFileString")
	
Case "CSV"
	for i = 0 to rs.Fields.count -1
			OutputStr = replace(rs.Fields(i).Name,","," ") 
			Response.write OutputStr & Separator
	next
	response.write CHR(10)
	Do while Not rs.EOF
		for i = 0 to rs.Fields.count -1
				OutputStr = ""
				Value = rs.Fields(i).Value
				' Clear CHR(13) and CHR(10)
				If Instr(Value,CHR(13)) > 0 then
					Value = Replace(Value,CHR(13)," ")
				End IF
				If Instr(Value,CHR(10)) > 0 then
					Value = Replace(Value,CHR(10),"")
				End IF
				If Not IsNull(Value) Then 
					OutputStr = rtrim(replace(Value,","," "))
				End if	
				Response.write OutputStr & Separator
		next
		response.write CHR(10)
	rs.moveNext
	Loop
Case "TAB"
	if Not(rs.EOF) then
			OutputStr = Null
			for i = 0 to rs.Fields.count -1
					OutputStr = rs.Fields(i).Name & Separator
					Response.write OutputStr
			next
			response.write CHR(10)
		Do while Not rs.EOF
			OutputStr = Null
			for i = 0 to rs.Fields.count -1
					OutputStr =  rs.Fields(i) & Separator
					Response.write OutputStr
			next
			response.write CHR(10)
		rs.moveNext
		Loop
	End If
Case "SPACE"
	if Not(rs.EOF) then
			OutputStr = Null
			for i = 0 to rs.Fields.count -1
					OutputStr = rs.Fields(i).Name & Separator
					Response.write OutputStr
			next
			response.write CHR(10)
		Do while Not rs.EOF
			OutputStr = Null
			for i = 0 to rs.Fields.count -1
					OutputStr =  rs.Fields(i) & Separator
					Response.write OutputStr
			next
			response.write CHR(10)
		rs.moveNext
		Loop
	End If
Case "COMMA"
	if Not(rs.EOF) then
			OutputStr = Null
			for i = 0 to rs.Fields.count -1
					OutputStr = chr(34) & rs.Fields(i).Name &  chr(34) & Separator
					Response.write OutputStr
			next
			response.write CHR(10)
		Do while Not rs.EOF
			OutputStr = Null
			for i = 0 to rs.Fields.count -1
					If isnumeric(rs.Fields(i)) Then
						OutputStr =  rs.Fields(i) & Separator
					else	
						OutputStr =  chr(34) & rs.Fields(i) & chr(34) & Separator
					end if
					Response.write OutputStr
			next
			response.write CHR(10)
		rs.moveNext
		Loop
	End If
CASE "FIXED"
	'Set dd = Conn.Execute("SELECT Table_Name, Table_Column_Name, Table_Column_Datatype FROM Data_Dictionary")
	if Not(rs.EOF) then
			OutputStr = ""
			for i = 0 to rs.Fields.count -1
					OutputStr = OutputStr &	padfield(rs.Fields(i).Name,rs.Fields(i).Name)
					Response.write OutputStr
			next
			response.write CHR(10)
		Do while Not rs.EOF
			OutputStr = ""
			for i = 0 to rs.Fields.count -1
					OutputStr = OutputStr &  padfield(rs.Fields(i).Name,rs.Fields(i).Value)
					Response.write OutputStr
			next
			response.write CHR(10) 
		rs.moveNext
		Loop
	End If
CASE "LABEL"
	if Not(rs.EOF) then
		LPad = "" 
		Do while Not rs.EOF
			L1C1 = ""
			L1C2 = ""
			L1C3 = ""
			L2C1 = ""
			L2C2 = ""
			L2C3 = ""
			L3C1 = ""
			L3C2 = ""
			L3C3 = ""
			L4C1 = ""
			L4C2 = ""
			L4C3 = ""
			L5C1 = ""
			L5C2 = ""
			L5C3 = ""
			L1C1Teachers = ""
			L1C2Teachers = ""
			L1C3Teachers = ""
			
			' First Recordset Row
			If Instr(Ucase(Session("SQLProj")),"NUMBER_OF_TEACHERS") Then
				L1C1Teachers = rs("Number_Of_Teachers")
			else
				L1C1Teachers = ""
			end if	
			for i = 0 to rs.Fields.count -1
					If Session("LabelName")="" then
						L1C1 =  Ucase(rs("Title") & " " & rs("First_Name") & " " & rs("Middle_Maiden_Name") & " " & rs("Last_Name") )
						L2C1 =  UCase(rs("Position_Type_Descr"))
					Else	
						L1C1 =  ""
						L2C1 =  Ucase(Session("LabelName") ) 
					End If
					If Session("PersonnelPath") <> "OTHER" then
						L3C1 =  Ucase(rs("Entity_Name"))
					Else
						L3C1 =  Ucase(rs("Location_Name"))
					End If	
					session("line4") = ""
					session("line5") = ""
					a = getaddress(Session("AddressType"))
					L4C1 =  session("line4")
					L5C1 =  session("line5")

			Next
			rs.MoveNext	
		  If Not rs.EOF Then
			' Second Row
			If Instr(Ucase(Session("SQLProj")),"NUMBER_OF_TEACHERS") Then
				L1C2Teachers = rs("Number_Of_Teachers")
			else
				L1C2Teachers = ""
			end if	
			for i = 0 to rs.Fields.count -1
					If Session("LabelName")="" then
						L1C2 =  UCase(rs("Title") & " " & rs("First_Name") & " " & rs("Middle_Maiden_Name") & " " & rs("Last_Name") )
						L2C2 =  UCASE(rs("Position_Type_Descr"))
					Else	
						L1C2 =  ""
						L2C2 =  Ucase(Session("LabelName") )
					End If
					If Session("PersonnelPath") <> "OTHER" then
						L3C2 =  Ucase(rs("Entity_Name"))
					Else
						L3C2 =  Ucase(rs("Location_Name"))
					End If	
					session("line4") = ""
					session("line5") = ""
					a = getaddress(Session("AddressType"))
					L4C2 =  session("line4")
					L5C2 =  session("line5")

			Next
			rs.MoveNext	
	     End If		
		  If Not rs.EOF Then
		   ' Third Row
			If Instr(Ucase(Session("SQLProj")),"NUMBER_OF_TEACHERS") Then
				L1C3Teachers = rs("Number_Of_Teachers")
			else
				L1C3Teachers = ""
			end if	
			for i = 0 to rs.Fields.count -1
					If Session("LabelName")="" then
						L1C3 =  UCASE(rs("Title") & " " & rs("First_Name") & " " & rs("Middle_Maiden_Name") & " " & rs("Last_Name") )
						L2C3 =  Ucase(rs("Position_Type_Descr"))
					Else	
						L1C3 =  ""
						L2C3 =  Ucase(Session("LabelName") ) 
					End If
					If Session("PersonnelPath") <> "OTHER" then
						L3C3 =  Ucase(rs("Entity_Name"))
					Else
						L3C3 =  Ucase(rs("Location_Name"))
					End If	
					session("line4") = "&nbsp;"
					session("line5") = "&nbsp;"
					a = getaddress(Session("AddressType"))
					L4C3 =  session("line4")
					L5C3 =  session("line5")

			Next
			rs.MoveNext	
		  End if
%>

<head>
</head>

<table width="100%"  height="8">
<B>
<% IF Session("LabelType") = "Pers" Then %>
<tr>
<td width="50" height="10"><font size="2"><b><font face="Times New Roman"><%=""%></font></b></font></td>
<td width="300" height="10"><font size="2"><b><font face="Times New Roman"><%=""%></font></b></font></td>
<td width="300" height="10"><font size="2"><b><font face="Times New Roman"><%=""%></font></b></font></td>
<td width="300" height="10"><font size="2"><b><font face="Times New Roman"><%=""%></font></b></font></td>
</tr>
<tr>
<td width="50" height="8"><font size="2"><b><font face="Times New Roman"><%=Lpad%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L1C1%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=L1C1Teachers%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L1C2%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=L1C2Teachers%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L1C3%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=L1C3Teachers%></font></b></font></td>
</tr>
<tr>
<td width="50" height="8"><font size="2"><b><font face="Times New Roman"><%=Lpad%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L2C1%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L2C2%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L2C3%></font></b></font></td>
</tr>
<tr>
<td width="50" height="8"><font size="2"><b><font face="Times New Roman"><%=Lpad%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L3C1%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=L1C1Teachers%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L3C2%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=L1C2Teachers%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L3C3%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=L1C3Teachers%></font></b></font></td>
</tr>
<% End if %>
<% IF Session("LabelType") = "Ent" Then %>
<tr>
<td width="50" height="8"><font size="2"><b><font face="Times New Roman"><%=Lpad%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L1C1%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L1C2%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L1C3%></font></b></font></td>
</tr>
<tr>
<td width="50" height="8"><font size="2"><b><font face="Times New Roman"><%=Lpad%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L2C1%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L2C2%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L2C3%></font></b></font></td>
</tr>
<tr>
<td width="50" height="8"><font size="2"><b><font face="Times New Roman"><%=Lpad%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L3C1%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=L1C1Teachers%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L3C2%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=L1C2Teachers%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L3C3%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=L1C3Teachers%></font></b></font></td>
</tr>
<tr>
<% End if %>
<td width="50" height="8"><font size="2"><b><font face="Times New Roman"><%=Lpad%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L4C1%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L4C2%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L4C3%></font></b></font></td>
</tr>
<tr>
<td width="50" height="8"><font size="2"><b><font face="Times New Roman"><%=Lpad%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L5C1%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L5C2%></font></b></font></td>
<td width="300" height="8"><font size="2"><b><font face="Times New Roman"><%=L5C3%></font></b></font></td>
</tr>
<tr>
<% 'End if %>
</B>
</table>
<%
		Loop
	End If

End Select
rs.close

Function getcolumnwidth(tblcoldatatype)
getparen = InStr(tblcoldatatype, "(")
If getparen = 0 Then
    getcolumnwidth = 10
Else
    getnumbers = Right(tblcoldatatype, Len(tblcoldatatype) - getparen)
    getcolumnwidth = Left(getnumbers, InStr(getnumbers, ")") - 1)
End If
End Function


Function padfield(fieldname,fieldvalue)
dd.MoveFirst
Do While NOT dd.EOF 
	If UCASE(dd("Table_Column_Name")) = UCASE(Fieldname) then
		columnwidth = getcolumnwidth(dd("TABLE_COLUMN_DATATYPE"))
		Exit Do
	End If
'response.write "DID IT"
dd.MoveNext
Loop
fldval = (fieldValue)
padfield = left(fldval & space(columnwidth),columnwidth)
end function


Function getaddress(addrtype)
If addrtype = "" then
	addrtype = "MAILING"
End IF	
addrtype = trim(ucase(addrtype))
Select Case addrtype
Case "MAILING"
	If rs("Mailing_Extend_Street_Address") <> "" then
		session("line4") =  Trim(Ucase(rs("Mailing_Extend_Street_Address")))
		session("line5") =  Trim(Ucase(rs("Mailing_City") & " " & rs("State") & " " & rs("Mailing_ZIP_Code")))
	else	
		session("line4") = Trim(Ucase(rs("Physical_Extend_Street_Address")))
		session("line5") = Trim(Ucase(rs("Physical_City") & " " & rs("State") & " " & rs("Physical_ZIP_Code")))
	end if	
Case "PHYSICAL"
		session("line4") = Trim(Ucase(rs("Physical_Extend_Street_Address")))
		session("line5") = Trim(Ucase(rs("Physical_City") & " " & rs("State") & " " & rs("Physical_ZIP_Code")))
Case "COURIER"
	If Len(Trim(rs("Courier_Number"))) > 0 then
		session("line4") =  Trim(Ucase(rs("Physical_City")))
		session("line5") =  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Courier&nbsp;#&nbsp;" & left(rs("Courier_Number"),2) & "-" & mid(rs("Courier_Number"),3,2) & "-" & right(rs("Courier_Number"),2)
	else
		If rs("Mailing_Extend_Street_Address") <> "" then
			session("line4") = Trim(Ucase(rs("Mailing_Extend_Street_Address")))
			session("line5") = Trim(Ucase(rs("Mailing_City") & " " & rs("State") & " " & rs("Mailing_ZIP_Code")))
		else	
			session("line4") = Trim(Ucase(rs("Physical_Extend_Street_Address")))
			session("line5") = Trim(Ucase(rs("Physical_City") & " " & rs("State") & " " & rs("Physical_ZIP_Code")))
		end if	
	end if
end select
getaddress = ""
end function

' ********************************************************************************
' ********************************************************************************
' Simple XML Library for Active Server Pages
' Copyright @ 2001
' Authors Ladislav Goc, Bruce Troutman

' This library helps with the data entry, save and manipulate of XML files.
' It is NOT written to be robust, but small, easy to use and modify. 

' You can use it freely, modify, add descriptions, functions as long
' as credit is given to the original authors.

'---------------------------------------------------------------------------
Function DeleteFile(strPath)

	' strPath must include physical Path
 
    On Error Resume Next

 	Dim objFSO, objFile
	set objFSO = CreateObject("Scripting.FileSystemObject")
	set objFile = objFSO.DeleteFile(strPath)

  If Err.Number = 0 Then
    objFile.Close
  End If
  DeleteFile = (Err.Number = 0)

End Function

' --------------------------------------------------------------------------------------------
Function UniqueID
' Author: Ladislav Goc, May 2001

' This function returns unique code, based on SERVER's date, time and user Session ID

' Unique ID has form :YYYYMMDD-HHMMSS-SSSSSSSSS
' it is string from date, time and session ID: SSSSSSSSS

StampDate=date()
StampTime=time
StampSession=session.sessionid

StampCharDate=datepart("yyyy",StampDate)&RIGHT("00"+TRIM(datepart("m",StampDate)),2)&RIGHT("00"+TRIM(datepart("d",StampDate)),2)
'RIGHT("00"+TRIM(datepart("d",StampDate)),2)
StampCharTime=RIGHT("00"+TRIM(CInt(hour(StampTime))),2)+RIGHT("00"+TRIM(CInt(minute(StampTime))),2)+RIGHT("00"+TRIM(CInt(Second(StampTime))),2)

Stamp=StampCharDate & "-" & StampCharTime & "-" & StampSession

UniqueID=Stamp

End Function 

'----------------------------------------------------------------------------------
Function XMLfromSQL(RootName, RootInfo, UniqueID)

' This is for a single flat file
' Assumes:
' that Session("SQLStr") exists
' that connection Conn exists
' that field names must Not have spaces in them

' SQLStr must have "As" clause removed
XMLSQLStr = Session("SQLStr")
Set rs = Conn.Execute(XMLSQLStr)
XMLSQLStr = EncodeXML(XMLSQLStr)

QUOT = Chr(34)  'double-quote character
RootName=Ucase(RootName)

XMLstr = ""
' The page header
XMLheader= "<?xml version=" & QUOT & "1.0" & QUOT & " encoding= " & QUOT & "ISO-8859-1" & QUOT & "?>" & vbCrlf 
XMLstr = XMLStr & XMLHeader

' The ROOT tag is called (RootName) and it has an atribut createdate 
' which shows, when has been (RootName) created and atribut UniqueID
' which is derived from actual Date and time

XMLRootElementOpen = vbCrlf & "<" & RootName & " " _
    & "CreateDate=" & QUOT & CDATE(now()) & QUOT & " " _
    & "RootInfo=" & QUOT & RootInfo & QUOT & " " _
    & "UniqueID=" & QUOT & UniqueID & QUOT & " " _
    & "SQL=" & QUOT & XMLSQLStr & QUOT & " >" & vbCrlf 
'response.write XMLRootElementOpen
'response.end
XMLstr = XMLStr & XMLRootElementOpen
Do while Not rs.EOF 
	XMLstr = XMLStr & "  <ROW>" &  vbCrlf
	for i = 0 to rs.Fields.count -1
		XMLElement =  ""
		FldName = Trim(rs.Fields(i).Name)
		FldValue = rs.Fields(i).Value
		If isnull(FldValue) then
			FldValue = ""
		End if	
		FldValue = Trim(CStr(FldValue))
	   	XMLElement = OpenTag(FldName) & EncodeXML(FldValue) &  CloseTag(FldName) & vbCrlf 
		XMLstr = XMLStr & XMLElement
	next
	XMLstr = XMLStr & "  </ROW>" &  vbCrlf
	rs.moveNext
Loop
XMLRootElementClose = vbCrlf & "</" & RootName & ">" 
XMLfromSQL = XMLStr & XMLRootElementClose
End Function

'--------------------------------------------------------------------------------------------
Function OpenTag(tagItem)
QUOT = Chr(34)  'double-quote character
OpenTag = "<" & Ucase(tagItem) & ">"
End Function

'--------------------------------------------------------------------------------------------
Function CloseTag(tagItem)
QUOT = Chr(34)  'double-quote character
CloseTag = "</" & Ucase(tagItem) & ">"
End Function

'--------------------------------------------------------------------------------------------
Function EncodeXML(String) 'This function replaces XML special characters
If len(String) > 0 Then
	String = Replace(String, "&", "&amp;")
	String = Replace(String,"<","&lt;")
	String = Replace(String,">","&gt;")
	String = Replace(String, Chr(34),"&quot;")
	String = Replace(String,"'","&apos;")
End If
EncodeXML = String
End Function

'--------------------------------------------------------------------------------------------
Function XMLMessage(String) 'This function returns messages in XML
sXML = "<?xml version=" & chr(34) & "1.0" & chr(34) & "?>" & VbCrLf
sXML = sXML & "<Root>" & VbCrLf
sXML = sXML & "<Message>" &  String & "</Message>" & VbCrLf
sXML = sXML & "</Root>"
Response.ContentType = "text/xml"
Response.Write sXML
If len(strTempFile) > 0 then
		oFile.Close
		DeleteFile(strTempFile)
		strTempFile = ""
End If
		
DeleteFile(strTempFile)
Response.end
End Function


%>