<% 

' Example, how to use XMLfromTable

' IF writeXMLFile(XMLfromTable("ENTRY1","LadiGoc"),"c:\LadiGoc.XML") then
'		response.write "O.K. XML created"
' Else
' 		response.write "Something is wrong. Probably bad rights for create in specified subdirectory."
' End If

'******************************************
' FUNCTIONS 
' *****************************************

Function UniqueID
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
Function XMLfromForm(RootName, UniqueID)

QUOT = Chr(34)  'double-quote character
RootName=Ucase(RootName)

' The page header
XMLheader= "<?xml version=" & QUOT & "1.0" & QUOT & "?>" & vbCrlf

' The ROOT tag is called (RootName) and it has an atribut createdate 
' which shows, when has been (RootName) created and atribut UniqueID
' which is derived from actual Date and time


strDate=CDATE(date())

strEntryForm = vbCrlf & "<" & RootName & " " _
    & "CreateDate=" & QUOT & strDate & QUOT & " " & "UniqueID=" & QUOT & UniqueID & QUOT &">" & vbCrlf 

strItem=""

FOR EACH element IN Request.form 
   strItem = strItem & OpenTag(element) & request.form(element) & CloseTag(element) & vbCrlf 
NEXT 

XMLfromForm=XMLheader & strEntryForm & strItem & CloseTag(RootName)

End Function

'----------------------------------------------------------------------------------
Function XMLCleanSQL(str)

' Removes "As" Clauses from str
Do While InStr(UCASE(str), " AS ") > 0
    str0 = str
    str0 = Left(str0, InStr(str0, " as "))
    str1 = Mid(str, InStr(str, "'") + 1)
    str1 = Mid(str1, InStr(str1, "'") + 1)
    str = str0 + str1
Loop
Session("XMLCleanStr") = str
XMLCleanSQL = str
End Function



'----------------------------------------------------------------------------------
Function XMLfromSQL(RootName, RootInfo, UniqueID)

' This is for a single flat file
' Assumes:
' that Session("SQLStr") exists
' that connection Conn exists
' that field names must Not have spaces in them

' SQLStr must have "As" clause removed
XMLSQLStr = XMLCleanSQL(Session("SQLStr"))
Set rs = Conn.Execute(XMLSQLStr)

QUOT = Chr(34)  'double-quote character
RootName=Ucase(RootName)

XMLstr = ""
' The page header
XMLheader= "<?xml version=" & QUOT & "1.0" & QUOT & "?>" & vbCrlf
XMLstr = XMLStr & XMLHeader

' The ROOT tag is called (RootName) and it has an atribut createdate 
' which shows, when has been (RootName) created and atribut UniqueID
' which is derived from actual Date and time

XMLRootElementOpen = vbCrlf & "<" & RootName & " " _
    & "CreateDate=" & QUOT & CDATE(date()) & QUOT & " " _
    & "RootInfo=" & QUOT & RootInfo & QUOT & " " _
    & "UniqueID=" & QUOT & UniqueID & QUOT & " " _
    & "SQL=" & QUOT & XMLSQLStr & QUOT &">" & vbCrlf 
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
	   	XMLElement = OpenTag(FldName) & "<![CDATA[" & FldValue & "]]>" & CloseTag(FldName) & vbCrlf 
		XMLstr = XMLStr & XMLElement
	next
	XMLstr = XMLStr & "  </ROW>" &  vbCrlf
	rs.moveNext
Loop
XMLRootElementClose = vbCrlf & "</" & RootName & ">" 
XMLfromSQL = XMLStr & XMLRootElementClose
End Function

'---------------------------------------------------------------------------------
Function WriteXMLFile(strContent,strXMLBookListFile)
  On Error Resume Next
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFSO.CreateTextFile(strXMLBookListFile, True)
  If Err.Number = 0 Then
    objFile.WriteLine strContent
    objFile.Close
  End If
  WriteXMLFile = (Err.Number = 0)
End Function

'-----------------------------------------------------------------------------------
Function OpenTag(tagItem)
QUOT = Chr(34)  'double-quote character
'IndentNumber = Cint(Indent) * 4
IndentStr = "    "
'For i = 1 to IndentNumber
'	Indentstr = Indentstr & Chr(32)
'Next	
OpenTag = Indentstr & "<" & Ucase(tagItem) & ">"
End Function

'-----------------------------------------------------------------------------------
Function CloseTag(tagItem)
QUOT = Chr(34)  'double-quote character
CloseTag = "</" & Ucase(tagItem) & ">"
End Function

%>
