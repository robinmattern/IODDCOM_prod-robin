<%@ LANGUAGE="VBScript" %>
<% Response.Buffer = True %>
<%
session("Sorttype")=""
session("Sortdirection")=""
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Logon</title>
</head>
<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->
<!--#INCL UDE FILE="_incheader.asp"-->
<!--#INCLUDE FILE="_incnav.asp"-->
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<!--#INCLUDE FILE="_incemail.asp"-->

<%
If len(trim(Session("SystemName"))) = 0 Then
	If Session ("SystemName") <> "CFO Measurement Tracking System" Then
		Session ("SystemName") = "CFO Measurement Tracking System"
		response.redirect "default.asp"
	End If
End If
%>

<script language=JavaScript runat=Server> 
function emailvalid(src) {
	emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[a-zA-Z]$"
	var regex = new RegExp(emailReg);
	return regex.test(src);	
}
</script> 

<% ' ================  Main Processing Begins Here  =======================

If ucase(request("DataAction")) = "REGISTER" then 
	session("dataaction") = "REGISTER"
	response.redirect "_logonreg.asp"
else
	session("dataaction") = ""
End if
If ucase(request("From")) = "LOGONINFO" then

Session("valPersonPassword") = Null
Session("valPasswordError") = Null
Session("valPersonPhone") = Null

		if emailcheck() and passwordcheck() then
			newpassword = trim(request("newpassword"))
			If LCase(newpassword) = "newuser" Then
				Session("valPersonPassword") = "* You must change your Password<br>to something other than 'newuser'"
				Session("valPasswordError") = "Password ERROR! (See Below)"
				response.redirect "_logoninfo.asp"
			End If
			If len(Trim(request("Phone"))) = 0 Then
				Session("valPersonPhone") = "You must enter a valid Phone Number"
				response.redirect "_logoninfo.asp"
			End If
	
			If trim(request("newemail")) <> trim(session("personemail")) then   'changed email
				newemail = True
				' Don't reverse right now bt 5/17/01  newpassword = ReverseString(newpassword)
			End If
			sqltext = "Update tperson set password = '" & newpassword &  "', PrefixName = '" & trim(request("PrefixName")) & "', FirstName = '" & trim(request("FirstName")) & "', LastName = '" & replace(trim(request("LastName")),"'","''") & "', SuffixName = '" & trim(request("SuffixName")) & "', Phone = '" & trim(request("phone")) & "', logon = '" & replace(trim(request("newemail")),"'","''") & "', changedat = '" & now() & "' where logon= '" & session("personemail") & "' and Password = '" & session("password") & "'"
			Conn.Execute sqltext
			If newemail then
				emailTo = request("newemail")
				emailSubject = "Your new password."
				emailMessage = "Your new password is:  " & newpassword  & chr(13) & chr(10) & "Please logon using this email address and this new password."
				SendEMail "fidosupport@8020data.com",emailto,emailsubject,emailmessage,""
				Session("message")="Your new password has been emailed to your new address.<br>Please logon with your new email address and system generated password<br>Click the checkbox on the logon screen to make changes."
				response.redirect "_message.asp"
			Else
				Session("message")="Your changes have been saved.<br>Please logon with your changed information."
				response.redirect "_message.asp"
			End if
		end if
End If

If ucase(request("From")) = "LOGONREG" then

	Session("valPersonPassword") = Null
	Session("valPasswordError") = Null
	Session("valPersonPhone") = Null
	
	orgcode = "DHS\" & request("OEProgram")
	newpassword = replace(left(time(),instr(time()," ")-1),":","")
	changedate = now()
	
	sqltext = "INSERT INTO tperson (role, custcode, organization, password, PrefixName, FirstName, LastName, SuffixName, Phone, Logon, createdby, createdat, changedby, changedat  ) VALUES ('REP','DHS-AAP', '" & orgcode & "', '" & newpassword &  "', '" & trim(request("PrefixName")) & "', '" & trim(request("FirstName")) & "', '" & replace(trim(request("LastName")),"'","''") & "', '" & trim(request("SuffixName")) & "', '" & trim(request("phone")) & "', '" & session("logon") & "', '" & session("logon") & "', '" & changedate & "', '" & session("logon")& "', '" & changedate & "')"
	'response.write "SQL: " & sqltext
	'response.end

	Conn.Execute sqltext
	emailTo = session("logon")
	emailSubject = "Your new password."
	emailMessage = "Your new password is:  " & newpassword  & chr(13) & chr(10) & "Please logon using this email address and this new password."
	SendEMail "fidosupport@fido.gov",emailto,emailsubject,emailmessage,""
	Session("message")="Your changes have been saved and your password has been emailed to your new address.<br>Please logon with your new email address and system generated password<br>Click the checkbox on the logon screen to make changes."
	response.redirect "_message.asp"
End If

if trim(request("logon"))="" then
%>
  <center>
  <table border="0" cellpadding="4" cellspacing="0">
  </table>
  </center>
<script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.Password.value == "")
  {
    alert("Please enter a value for the \"Password\" field.");
    theForm.Password.focus();
    return (false);
  }
  return (true);
}
//--></script>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.logon.value == "")
  {
    alert("Please enter a value for the \"Logon\" field.");
    theForm.logon.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="_logon.asp" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
<input type="hidden" value="<%=strCaller%>" name="Caller">


<div align="left">
<table border = "0" cellpadding="0" bordercolor="#003399" cellspacing="0" style="border-collapse: collapse">
  <tr><td align="center" width="890" bgcolor="#FFFFFF">
  <p align="left">
  <font color="#800000" size="5"><b>Logon</b> </font>
	<p align="left">
  <b>
	<a target="_blank" title="Plays web movie" href="http://fido.gov/FidoLogon.html">
	How to Logon to a Fido system</a></b></tr>
  <tr><td align="center" width="890" bgcolor="#FFFFFF">
  <div align="center">
    <div align="left">
    <table border="0" cellpadding="5" cellspacing="0" width="452">
      <tr>
      <td bgcolor="#FFFFFF" align="left" width="104">
        <p align="left">
        <b><font color="#000080">Please enter your </font></b>
        <font color="#000080"><b>Email</b></font>
        </p>
      </td>
    <center>
      <td bgcolor="#FFFFFF" width="324" align="left">
        <!--webbot bot="Validation" s-display-name="Logon" b-value-required="TRUE" --><input name="logon" size="30">
      </td>
      </tr>
    </center>
      <tr>
      <td bgcolor="#FFFFFF" align="left" width="104">
        <p align="left"><b><font color="#000080">Please enter your Password</font></b></p>
      </td>
    <center>
      <td bgcolor="#FFFFFF" width="324">
        <p align="left"><input type="password" name="Password" size="15"></p>
      </td>
      </tr>
      <tr>
      <td colspan="2" bgcolor="#FFFFFF" width="414">
        <div align="left">
          <table border="0" cellpadding="2" cellspacing="0">
            <tr>
              <td><input type="checkbox" name="chkForgotPassword" value="ON">  <i>
				<font color="#003399"><b>I forgot my password. Please send
                it to me.</b></font></i></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFFF"><input type="checkbox" name="chkChangeInfo" value="ON"> 
				<i> <font color="#003399"><b>I
                want to change my personal information.</b></font></i></td>
            </tr>
    <tr>
      <td colspan="2" align="center" bgcolor="#FFFFFF">
        <input type="submit" value="Continue" name="btn">
        <br>
        
        </td>
    </tr>
</table>
        </div>
      </tr>
    </table>
    </div>
    </center>
  </div>
  </tr>
  <center>
  </center>
</table>
</div>
  <center>
  <p></p>
  </form>
  </center>


</body>
</html>
<!--#INCLUDE FILE="_incfooter.asp"-->
<%
else
	Session("logon") = request("logon")
	If CheckLogon() Then
		' Show appropriate Form based on Logon Privileges
		' Any Data Entry has Write Permission
		' POC = Point of Contact, REP = Reporter
		Session("FY") = Session("CurrentFY")
		SESSION("Message")=""
		Session("AgencyBureauCode") = ""
		Session("BureauCode") = ""
		Session("ReadWrite") = "Read"
		Session("agencyID") = ""
		Session("DataEntryRole") = ""
		Session("POCEmail") = ""
		Session("AgencyAbbr") = ""

		If instr(session("grouplist"),"REP") or instr(session("grouplist"),"POC") then
			Session("ReadWrite") = "Write"
			If instr(session("grouplist"),"POC") then
				Session("DataEntryRole") = "POC"
					
			Else
				Session("DataEntryRole") = "REP"
			End If
'			SQLStr = "Select AgencyID, AgencyAbbr, AgencyName from tAgency WHERE AgencyAbbr = '" & SESSION("PersonOrganization") & "'"
'			set rs = Conn.Execute(SQLStr)
'			Session("agencyID") = rs("agencyID")
'			Session("agencyAbbr") = rs("agencyAbbr")
'			Session("agencyName") = rs("agencyName")
'			If Session("DataEntryRole") = "REP" then ' get POC email
'				SQLStr = "select logon from tperson where role = 'POC' and custcode = '" & session("mtstype") & "' and organization = '" & session("agencyabbr") & "'"
'				session("db") = sqlstr
'				set rs1 = Conn.Execute(SQLStr)
'				Session("POCEmail") = ""
'				do while not rs1.eof
'					Session("POCEmail") = Session("POCEmail") & ", " & rs1("Logon")
'					rs1.movenext
'				loop	
'			End IF					
			response.redirect session("MainTableCall")
		End If			
		If Instr(Session("PersonOrganization"),"/RO") > 0 and Session("PersonOrganization") <> session("mtstype") & "/RO" then  'type or Agency ReadOnly
				AgencyAbbr = left(Session("PersonOrganization"),Instr(Session("PersonOrganization"),"/RO")-1)
				SQLStr = "Select AgencyID, AgencyAbbr, AgencyName from tAgency WHERE AgencyAbbr = '" & AgencyAbbr & "'"
				set rs = Conn.Execute(SQLStr)
				Session("agencyID") = rs("agencyID")
				Session("agencyAbbr") = rs("agencyAbbr")
				Session("agencyName") = rs("agencyName")
				response.redirect "_notready.asp"
		End if
		If Instr(ucase(session("grouplist")),"ADM") > 0 Then 'If instr(session("grouplist"),"Administrators") then
			Session("ReadWrite") = "Write"
		End If	
		response.redirect session("MainTableCall")
	End if
End If

'=============================================
'Log On Functions
'=============================================
Function checklogon()
CheckLogon = False
IF request("chkForgotPassword")="ON" or len(trim(Request("password"))) = 0 then
	SQLStr = "SELECT * FROM tPerson WHERE tPerson.Logon= '" & replace(Request("logon"),"'","''") & "'"
	Set rs = Conn.Execute(SQLStr) 
%>
<!-- Perform Logon check only -->
	<%If request("chkForgotPassword")<>"ON" and rs.EOF Then %>
		<br>
		<br>
		
		<p align="left"><br>
		<font color="#000080">
		<strong>Your email address has not been registered in the system.</strong></font></p>
		<p align="left"><font color="#000080">
		<strong>Please click the Register button </strong> </font> </p>
		<p align="left"><font color="#000080">
		<strong>or</strong><br>
		<br>
		<strong>Click the Return to Logon button, if you entered your email address 
		incorrectly.</strong> </font> </p>
		
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
		  <tr>
		    <td align="center" width="100%" valign="middle"><form method="POST" Action="_logon.asp">
			<input type="hidden" value="<%=strCaller%>" name="Caller">
		      <p align="left"><font color="#400040">
				<input type="submit" value="Register" name="DataAction">&nbsp;&nbsp; <input type="submit" value="Return to Logon" name="DataAction"></font></p>
		    </form>
		    </td>
		  </tr>
		</table>
		<p>
		<!--#INCLUDE FILE="_incfooter.asp"-->
		
		<% rs.Close
		SET rs = Nothing%> 
		<%Exit Function%> 
		<!--#INCLUDE FILE="_incfooter.asp"-->
		
		<% rs.Close
			SET rs = Nothing
			Exit Function 
	Else
		IF request("chkForgotPassword")="ON" then
			sfrom = "fidosupport@fido.gov"
			sto = rs("logon")
			ssubj = "Your Password"
			smessage = "Your password is: " & rs("password")
			SendEMail sfrom, sto , ssubj, smessage, ""
			session("message") = "Your password has been sent to your email address."
			response.redirect "_message.asp"
		Else
%>
<br>
<br>

<p align="left"><br><font color="#000080">
<strong>Your password is blank.</strong></font></p>
<p align="left">
<br>
<strong>Click the Return to Logon button</strong> </font> </p>
<p align="left"> </p>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td align="center" width="100%" valign="middle"><form method="POST" Action="_logon.asp">
	<input type="hidden" value="<%=strCaller%>" name="Caller">
      <p align="left"><font color="#400040">
		<input type="submit" value="Return to Logon" name="DataAction"></font></p>
    </form>
    </td>
  </tr>
</table>
<p>
<!--#INCLUDE FILE="_incfooter.asp"-->
<%
		End if
	End if	
End If
 
%>
<!-- Perform Logon and Password check -->
<% 
SQLStr = "SELECT * FROM tPerson WHERE tPerson.Logon= '" & replace(Request("logon"),"'","''") & "' AND tPerson.Password='" & replace(Request("password"),"'","''") & "'"
'response.write SQLStr
'response.end

Set rs = Conn.Execute(SQLStr) 

If rs.EOF Then %>
<br>
<br>

<p align="left"><br><font color="#000080">
<strong>Your email address and password are not found in the system.</strong></font></p>
<p align="left">
<br>
<strong>Click the Return to Logon button</strong> </font> </p>
<p align="left"> </p>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td align="center" width="100%" valign="middle"><form method="POST" Action="_logon.asp">
	<input type="hidden" value="<%=strCaller%>" name="Caller">
      <p align="left"><font color="#400040">
		<input type="submit" value="Register" name="DataAction">&nbsp;&nbsp; <input type="submit" value="Return to Logon" name="DataAction"></font></p>
    </form>
    </td>
  </tr>
</table>
<p>
<!--#INCLUDE FILE="_incfooter.asp"-->

<% rs.Close
SET rs = Nothing%> 
<%Exit Function%> 
<% End If %> 

<!-- Perform Active User check --> 
<%If rs("ActiveUser")<>"Yes"  Then %> </p>

<p align="left"><br>
<font color="#000080">
<strong>Your Account is not active. Please contact your POC to have your user information
updated.</strong><br>
<br>
<strong>Please click the Return to Logon button.</strong> </font> </p>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td align="center" width="100%" valign="middle"><form method="POST" Action="default.asp">
		<input type="hidden" value="<%=strCaller%>" name="Caller">
      <p align="left"><font color="#400040"><input type="submit" value="Return to Logon" name="DataAction"></font></p>
    </form>
    </td>
  </tr>
</table>

<p>
<!--#INCLUDE FILE="_incfooter.asp"-->
	<% rs.Close
	SET rs = Nothing%><%Exit Function%><% End If %><%
'Get XML Info
Session("OKforXML") = "No"
strUA = Request.ServerVariables("HTTP_USER_AGENT")
if instr(strUA, "MSIE") then
	lnIntVer = CInt(Mid(strUA, InStr(strUA, "MSIE") + 5, 1))
	if lnIntVer >= 5 then
		Session("OKforXML") = "Yes"
	else
	 	'Response.Write("Please upgrade at <a href='http://www.microsoft.com/ie/'>http:///www.microsoft.com/ie/</a>")
	end if
else
	'Response.Write("Not IE<br>")
end if
'Get Browser Info
Set bc = Server.CreateObject("MSWC.BrowserType")   
SESSION("Browser") = bc.browser & " " & bc.version
SESSION("PersonID")=rs("PersonID")
SESSION("Password")=request("Password")
session("grouplist") = rs("Role")
session("role") = rs("Role")
SESSION("PersonPrefixName") = rs("PrefixName")
SESSION("PersonFirstName") = rs("FirstName")
SESSION("PersonLastName") = rs("LastName")
SESSION("PersonSuffixName") = rs("SuffixName")
SESSION("PersonName") = Trim(rs("PrefixName") & " " & rs("FirstName") & " " & rs("LastName") & " " & rs("SuffixName"))
SESSION("PersonLogon")=rs("Logon")
SESSION("PersonEmail")=SESSION("PersonLogon")
SESSION("PersonPhone")=rs("Phone")
SESSION("Organization")= rs("Organization")
SESSION("PersonPhone")=rs("Phone")
SESSION("PersonOrganization")=rs("Organization")

SESSION("ChangedBy") = trim(rs("Firstname")& " " & rs("Lastname")& " - " &rs("logon")& ", " & rs("Phone")) 

IF lcase(Session("Password"))="newuser" then
	response.redirect "_logoninfo.asp"
End if	


SQLStr="Update tPerson SET LASTLOGONDATE = GETDATE(), BROWSER =  ' " & SESSION("BROWSER") & " ' WHERE PersonID = " & SESSION("PersonID")
SET rs = Conn.Execute(SQLStr)
SET rs = Nothing



' Must validate logon and password first, then
IF request("chkChangeInfo")="ON" then
	response.redirect "_logoninfo.asp"
End if	

CheckLogon = True

End FUNCTION%><!-- System functions --><%
function reversestring(astring)
reversestring = ""
for i = 1 to len(astring)
	reversestring = mid(astring, i,1) + reversestring
next 
end function


function emailcheck()
If len(trim(request("newemail")))>0 Then 
	If trim(request("newemail")) <> trim(request("confirmemail")) then 
		session("message") = "Your email entries do not match."
		response.redirect "_message.asp"
	End If
	If not emailvalid(request("newemail")) then 
		session("message") = "Your email entries are not valid email addresses."
		response.redirect "_message.asp"
	End If
End If
emailcheck = True
End Function

function passwordcheck()
If len(trim(request("newpassword")))>0 Then 
	If trim(request("newpassword")) <> trim(request("confirmpassword")) then 
		session("message") = "Your password entries do not match."
		response.redirect "_message.asp"
	End If
End If
passwordcheck = True
End Function
%>