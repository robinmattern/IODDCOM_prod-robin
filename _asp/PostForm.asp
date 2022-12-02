<%@ LANGUAGE="VBSCRIPT" 
  'To Do: Still need to take into account if the user has put a CRLF at the end of each line.
%>
<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">

<title>Document Title</title>
<meta name="Microsoft Theme" content="none, default">
</head>

<body>
<a HREF="../forum.asp">

<p align="center"><img BORDER="0" SRC="../BackToTheForum.gif" WIDTH="221" HEIGHT="17"></a> </p>

<hr>
<%
  intID = Request.QueryString ("ID")
  If intID > 0 Then 'Value obtained.
    SQLQuery = "SELECT DISTINCTROW Message.ID, Message.Subject, " _
             & "Message.From, Message.Email, Message.Body, Message.When, " _
             & "Message.MsgLevel, Message.PrevRef FROM Message " _
             & "WHERE (((Message.ID)= " & intID & "));"
    Set RS= conn.Execute(SQLQuery)
    strSubjectText = RS.Fields("Subject")
    If UCase(Left(strSubjectText,3)) <> "RE:" Then 'Add reply indicator.
      strSubjectText = "Re: " & strSubjectText
    End If
    lngID = RS.Fields("ID")
    intMsgLevel = RS.Fields("MsgLevel")
    lngPrevRef = RS.Fields("PrevRef")
  End If
%>

<form NAME="frmMessage" METHOD="POST" ACTION="../DoPost.asp">
  <input type="hidden" name="hidID" value="<%= lngID %>"><input type="hidden" name="hidMsgLevel" value="<%= intMsgLevel %>"><input type="hidden" name="hidPrevRef" value="<%= lngPrevRef %>"><table WIDTH="590" BORDER="0" CELLSPACING="0" CELLPADDING="4">
    <tr>
      <td ALIGN="RIGHT" VALIGN="TOP"><font FACE="ARIAL"><b>Name:<br>
      </b></font></td>
      <td ALIGN="LEFT" VALIGN="TOP"><input TYPE="TEXT" NAME="txtName" VALUE SIZE="50"><br>
      </td>
    </tr>
    <tr>
      <td ALIGN="RIGHT" VALIGN="TOP"><font FACE="ARIAL"><b>Email:<br>
      </b></font></td>
      <td ALIGN="LEFT" VALIGN="TOP"><input TYPE="TEXT" NAME="txtEmail" VALUE SIZE="50"><br>
      </td>
    </tr>
    <tr>
      <td ALIGN="RIGHT" VALIGN="TOP"><font FACE="ARIAL"><b>Subject:<br>
      </b></font></td>
      <td ALIGN="LEFT" VALIGN="TOP"><input TYPE="TEXT" NAME="txtSubject" VALUE="<% = strSubjectText %>" SIZE="50"><br>
      </td>
    </tr>
    <tr>
      <td COLSPAN="2" ALIGN="CENTER" VALIGN="TOP"><font FACE="ARIAL"><b>Message Text<br>
<%
  If intID > 0 Then
    datWhen = RS.Fields("When")
    strFrom = RS.Fields("From")
    strBodyIn = RS.Fields("Body") 'Need to process it for reply format.
    strBody = GenerateBody(datWhen, strFrom, strBodyIn, 78)
  Else
    datWhen = ""
    strFrom = ""
    strBody = ""
  End If 
%>      <textarea NAME="txtBody" ROWS="15" COLS="78"><% = strBody %> <!-- Nec to avoid CRLF -->
</textarea><br>
      </b></font></td>
    </tr>
    <tr>
      <td COLSPAN="2" ALIGN="CENTER"><% If intID > 0 Then 'Post Reply %> <input TYPE="image" VALUE="Submit Post Reply" SRC="PostReply.gif" WIDTH="122" HEIGHT="22"> <% Else 'Post Message %> <input TYPE="image" VALUE="Post Message" SRC="PostMessage.gif" WIDTH="157" HEIGHT="22"> <% End If %> </td>
    </tr>
  </table>
</form>
</body>
</html>
<%
  Function GenerateBody(datWhen, strFrom, strOrigBody, intWidth)
  'Still need to take into account if the user has put a CRLF at the end of each line.
  
    Dim strCopy, strCarryOver, strOutput, strProposedBodyLine
    Dim intCopyLength, intLastSpace, intCharPos, intLenProposed
    
    strCopy = strOrigBody
    intCopyLength = Len(strCopy)
    intLastSpace = -1 'No last space (yet).
    intCharPos = 1
    strCarryOver = ""
    strOutput = ""
    While intCopyLength > intWidth + 2 '+ 2 Because of prepending '> ' to each body line.
      strProposedBodyLine = strCarryOver & Left(strCopy, intWidth - Len(strCarryOver) - 2)
      intLenProposed = Len(strProposedBodyLine)
      strCopy = Right(strCopy, Len(strCopy) - intLenProposed + Len(strCarryOver)) 'May need to + 1
      intCopyLength = Len(strCopy)
      intLastSpace = LastPositionIn(strProposedBodyLine, " ")
      If intLastSpace > 0 Then 'Space found in body line.
        strCarryOver = Right(strProposedBodyLine, intLenProposed - intLastSpace)
        strProposedBodyLine = "> " & Left(strProposedBodyLine, intLastSpace) & vbCrLf 'Still to do: Take one off for space.
      Else
        strCarryOver = "" 'Reset carry over.
        strProposedBodyLine = "> " & strProposedBodyLine & vbCrLf
      End If
      strOutput = strOutput & strProposedBodyLine
    Wend
    strOutput = strOutput & "> " & strCarryOver & strCopy
    GenerateBody = "On " & datWhen & ", " & strFrom & " wrote: " & vbCrLf _
                 & strOutput
  End Function
  
  Function LastPositionIn(strSource, strSearchChar)
    
    Dim strCopy
    Dim intCharPos1, intCharPos2
    
    strCopy = strSource
    intCharPos1 = 0
    intCharPos2 = InStr(strSource, strSearchChar)
    While intCharPos2 > 0
      intCharPos2 = InStr(strCopy, strSearchChar)
      strCopy = Right(strCopy, Len(strCopy) - intCharPos2)
      intCharPos1 = intCharPos1 + intCharPos2
    Wend
    LastPositionIn = intCharPos1
    
  End Function
  
%>
