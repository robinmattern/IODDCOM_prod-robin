<%
'******************************************************
Function ReadConfiguration()
Set rs = Conn.Execute("Select * from tConfiguration")
do while not rs.eof
	Select Case rs("Description")
	Case "VisitorCount"
		' Increment Counter
		Session("Visitor") = CLng(rs("Settings"))
		X = CLNG(rs("Settings")) + 1
		sqltext = "UPDATE tConfiguration SET Settings = '" & X & "' WHERE  Description = 'VisitorCount'"	
		Conn.Execute(sqltext)
	End Select
	Session(rs("description")) = rs("Settings")
	rs.movenext
loop		
rs.Close
SET rs = Nothing
End Function
%>