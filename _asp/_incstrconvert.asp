<%


Function strconvert(convertstring, converttype)

' Usage: strconvert("stuart.troutman@8020data.com","html")

' Check for Nulls
If isnull(convertstring) then
	strconvert = "Nulls are not allowed"
	exit Function
End if	

ctype = lcase(converttype)
cstrin  = convertstring
cstrout = ""

' Check for bad ctype
If instr("html,ascii,password",ctype) = 0 then
	strconvert = "Improper Type: " & converttype
	Exit Function
End If

'Proceed based on type

Select Case ctype
Case "html" 
	' Do html stuff
	For I = 1 to len(cstrin) 
'		cstrout = cstrout & mid(cstrin,I,1) & ": " & "&#" & asc(mid(cstrin,I,1)) & ";" & " "
		cstrout = cstrout & "&#" & asc(mid(cstrin,I,1)) & ";"
	Next		 
Case "ascii" 
	' Do ascii stuff
	For I = 1 to len(cstrin) 
'		cstrout = cstrout & mid(cstrin,I,1) & ": " & right("0" & asc(mid(cstrin,I,1)),3) & " "
		cstrout = cstrout & right("0" & asc(mid(cstrin,I,1)),3)
	Next		

Case "password"
	' Do password stuff
	For I = 1 to len(cstrin) 
        If int(asc(mid(cstrin,I,1))) Mod 2 = 0 then   'it's even so add 3
        '	cstrout = cstrout & mid(cstrin,I,1) & ": " & right("0" & int(asc(mid(cstrin,I,1))+3),3) & " "
        	cstrout = cstrout & right("0" & int(asc(mid(cstrin,I,1))+3),3)
        else  ' it's odd subtract 3
        '	cstrout = cstrout & mid(cstrin,I,1) & ": " & right("0" & int(asc(mid(cstrin,I,1))-3),3) & " "
        	cstrout = cstrout & right("0" & int(asc(mid(cstrin,I,1))-3),3)
        end if
	Next		
End Select

strconvert = cstrout 

End Function


%>

