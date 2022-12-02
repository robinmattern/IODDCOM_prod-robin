<%@ LANGUAGE="VBSCRIPT" 
%>
<!--#INCLUDE FILE="inccreateconnection.asp"-->
<html>

<head>
<meta NAME="GENERATOR" CONTENT="Microsoft FrontPage 6.0">

<title>Messages Application</title>
<meta name="Microsoft Theme" content="none, default">
</head>

<%
  strQuote = Chr(34)
  Const cMaxTopLevel = 20
  Const cID = 0
  Const cSubject = 1
  Const cFrom = 2
  Const cEmail = 3
  Const cWhen = 4
  Const cMsgLevel = 5
  Const cPrevRef = 6
  Const cThreadPos = 7
%>

<h1><font color="#FFFF00">Welcome to the Forum</font></h1>

<p>Message List: </p>

<hr>
<%
  'Get the top level items, and sort in decending order.
  SQLQuery = "SELECT Message.ID, Message.Subject, Message.From, " _
           & "Message.Email, Message.When, Message.MsgLevel, " _
           & "Message.PrevRef, Message.ThreadPos FROM Message " _
           & "WHERE (((Message.MsgLevel)=1)) ORDER BY Message.When DESC"
response.write sqlquery 
 Set RS = conn.Execute(SQLQuery)
  If Not RS.EOF Then
    arrRecTop = RS.GetRows(cMaxTopLevel)
    
    SQLQuery = "SELECT Message.ID, Message.Subject, Message.From, Message.Email, " _
             & "Message.When, Message.MsgLevel, Message.PrevRef, " _
             & "Message.ThreadPos FROM Message WHERE (((Message.MsgLevel)>1))"
    
    Set RS=Conn.Execute(SQLQuery)
    If Not RS.EOF Then
      arrRecRest = RS.GetRows()
    Else
      Dim arrRecRest
      arrRecRest = Empty
    End If
    
    gLastMsgLevel = 0
    For intTopLevelRow = 0 To UBound(arrRecTop, 2) 'Iterate through all top rows.
      GenList arrRecTop(cMsgLevel, intTopLevelRow), arrRecTop, intTopLevelRow
      If Not IsEmpty(arrRecRest) Then
        ExpandFrom arrRecTop(cID, intTopLevelRow), 1
      End If
    Next
    GenList 0, 0, 0
  Else
    Response.Write "<B>No messages<B>"
  End If
%>
<%
  Sub ExpandFrom(lngID, intThreadPos)
    If intThreadPos <= 10 Then 'Continue processing.
      For lngRow = 0 To UBound(arrRecRest, 2)
        If (arrRecRest(cPrevRef, lngRow) = lngID) And (arrRecRest(cThreadPos, lngRow) = intThreadPos) Then 'A child message found.
          'Output row.
          GenList arrRecRest(cMsgLevel, lngRow), arrRecRest, lngRow 'Expand current branch.
          ExpandFrom arrRecRest(cID, lngRow), 1
          Exit For
        End If
      Next
      ExpandFrom lngID, intThreadPos + 1  'Expand branches below current branch.
    End If
  End Sub
%>

<hr>
<a HREF="PostForm.asp">

<p align="center"><img BORDER="0" SRC="PostNewMessage.gif" WIDTH="213" HEIGHT="22"></a> </p>
</html>
<%
  Function GenList(intNewMsgLevel, arrSourceArray, intRow)
    Dim intIndex
    
    For intIndex = gLastMsgLevel To intNewMsgLevel -1 'Upping the list levels in the iterations.
      Response.Write "<UL TYPE=DISC>" & vbCrlf
    Next
    For intIndex = intNewMsgLevel To gLastMsgLevel -1' Downing the list levels in the iterations.
      Response.Write "</UL>"
    Next
    If intNewMsgLevel > 0 Then
      Response.Write ListItem(arrSourceArray, intRow)
    End If
    gLastMsgLevel = intNewMsgLevel
  End Function
  
  Function ListItem(arrSourceArray, intRow)
    ListItem = "<LI>" & vbCrlf & "<A HREF=" & strQuote & "GetMessage.asp?ID=" _
             & arrSourceArray(cID, intRow) & strQuote & "> " _
             & arrSourceArray(cSubject, intRow) & "</A> - <B>" & From (arrSourceArray,intRow) _
             & "</B> <I>" & arrSourceArray(cWhen, intRow) & "</I>" & strEnd & vbCrlf
  End Function
  
  Function From(arrSourceArray, intRow)
    strName = arrSourceArray(cFrom, intRow)
    strEmail = arrSourceArray(cEmail, intRow)
    If InStr(strEmail, "@") > 0 Then 'Email parameter given and is OK.
      From = "<A HREF=" & strQuote & "mailto:" & strEmail & strQuote & ">" & strName & "</A>"
    Else
      From = strName
    End If
  End Function
%>