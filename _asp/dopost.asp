<%@ LANGUAGE="VBSCRIPT" %>

<!-- # include "adovbs.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">

<title>Document Title</title>
<meta name="Microsoft Theme" content="none, default">
</head>
<body>
<p>


<%
  Const cMaxMessageLevel = 10
  
  Dim rsAddMessage
  
  lngID = Request.Form("hidID")
  'Get common data.
  strFrom = Request.Form("txtName")
  strEmail = Request.Form("txtEmail")
  strSubject = Request.Form("txtSubject") 
  strBody = Request.Form("txtBody")
  lngPrevRef = Request.Form("hidID")
  
'  Set Conn = Session("IO")
  If lngPrevRef <> "" Then 'Post reply.
    intMsgLevel = Request.Form("hidMsgLevel")
    If intMsgLevel = cMaxMessageLevel Then 'Max message level reached.
      intNewMsgLevel = cMaxMessageLevel
    Else
      intNewMsgLevel = intMsgLevel + 1
    End If
    SQLQuery = "SELECT IIf(Max([ThreadPos])=NULL,1,Max([ThreadPos])+1) " _
             & "AS NewThreadPos FROM Message " _
             & "WHERE (((Message.PrevRef)= " & lngID & "));"  
    Set rsThreadPos = conn.Execute(SQLQuery)
    intNewThreadPos = rsThreadPos.Fields("NewThreadPos") 
    
  Else 'Post new message.
    Response.Write "Post new message."
    intNewMsgLevel = 1
    lngPrevRef = 0
    intNewThreadPos = 1
  End If
  
  Set rsAddMessage = Server.CreateObject("ADODB.Recordset")
  rsAddMessage.Open "Message", Conn, 1, 4
  Application.Lock
  rsAddMessage.AddNew
  rsAddMessage.Fields("From") = strFrom
  rsAddMessage.Fields("Email") = strEmail
  rsAddMessage.Fields("Subject") = strSubject
  rsAddMessage.Fields("Body") = strBody
  rsAddMessage.Fields("When") = CStr(Now())
  rsAddMessage.Fields("MsgLevel") = intNewMsgLevel
  rsAddMessage.Fields("PrevRef") = lngPrevRef
  rsAddMessage.Fields("ThreadPos") = intNewThreadPos
  rsAddMessage.UpdateBatch
  Application.Unlock
  rsAddMessage.Close
  Set rsAddMessage = Nothing
%>
<p>

<hr>
<center>
<h1> Message Received </h1>
<b><a HREF="forum.asp">Click here to return to the Forum</a></b>
</center><p>
<hr>
</body>
</html>
