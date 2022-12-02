<%
' Must change these for each new environment
servername = "localhost"
database = "io"
username = "io"
password = "passio1"

SET Conn = SERVER.CREATEOBJECT("ADODB.Connection")
' DSN-Less OLEDB SQLServer
Conn.Open "PROVIDER=MSDASQL;DRIVER={SQL Server};Server=" & servername & ";Database=" & database & ";Uid=" & username & ";Pwd=" & password

' DSN-Less OLEDB Access MDB
'CD.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\inetpub\secureaddress\mdb\TEST.mdb;User ID=;"

Function CleanSQL(SQLStr)

Syspos = InStr(SQLStr, "sysob")  'Get sys reference position - Sysobjects
Unionpos = InStr(SQLStr, "union")  'Get union position
'Deletepos = InStr(SQLStr, "delete")  'Get union position
'Updatepos = InStr(SQLStr, "update")  'Get union position
if SysPos > 1 then
	session("message") = SQLStr & "<BR> SQL Injection attempt. Last request aborted."
	response.redirect "_message.asp"
end if

Semipos = InStr(SQLStr, ";")  'Get semi-colon position
OQuotepos = InStr(SQLStr, "'") ' Get Open Quote position
AtAtPos = InStr(SQLStr, "@@") ' Get @@ position
DashDashPos = InStr(SQLStr, "--") ' Get -- position

If AtAtPos > 1 then
	SQLStr = Left(SQLStr, AtAtPos - 2)
End If

If DashDash > 1 then
	SQLStr = Left(SQLStr, DashDashPos - 2)
End If

' If semi-colons
If Semipos > 0 Then
    ' If no quotes
    If OQuotepos = 0 Then
        SQLStr = Left(SQLStr, Semipos - 1)
    Else
        'If semi-colon is before quote
        If Semipos < OQuotepos Then
            SQLStr = Left(SQLStr, Semipos - 1)
        Else
            Do While OQuotepos > 0 'get Close Quote
                CQuotepos = InStr(OQuotepos + 1, SQLStr, "'")
                If Semipos < CQuotepos Then 'semi is between quotes
                    Do While Semipos > OQuotepos And Semipos < CQuotepos ' check for next semi-colon
                        Semipos = InStr(Semipos + 1, SQLStr, ";")
                    Loop
                End If
                If Semipos = 0 Then  'No semis outside of quotes
                    OQuotepos = 0
                Else
                    OQuotepos = InStr(CQuotepos + 1, SQLStr, "'")
                End If
            Loop
        End If
        If Semipos > 0 Then
            SQLStr = Left(SQLStr, Semipos - 1)
        End If
    End If
End If
	
CleanSql = SQLStr
End Function
%>