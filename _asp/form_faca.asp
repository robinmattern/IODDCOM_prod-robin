<%@ LANGUAGE="vbscript" %> <% Response.Buffer = True %>
<!--#INCLUDE FILE="incExpires.asp"-->
<!--#INCLUDE FILE="incLogonCheck.asp"-->
<%
' Usage:
' For List:   Call RootFileName & ""
' For Form:   Call RootFileName & "?ID=xxxx"  where xxxx is the strTableName & "ID"
' For Add:    Call RootFileName & "?ID=0      where 0 means add
' For Sort:   Call RootFileName & "?Sort=aaaaaaa"  where aaaaaaa is the Field Name to sort on
' For Action: Set Buttons in the Form or List
' For Mark Complete. Set session("Process") = "MARK COMPLETE"

' Requires the following Include files:
' ----------------------------------------
'    incExpires.asp
'    incLogonCheck.asp
'    inccommonfunctions.asp
'    inccreateconnection.asp
'    incheader.asp
'    incnav.asp
'    incbodyline.asp
'    incfooter.asp
'    incmessage.asp
'    incgeneratereport.asp
'    incgeneratefile.asp
%>
<!--#INCLUDE FILE="inccommonfunctions.asp"-->
<!--#INCLUDE FILE="inccreateconnection.asp"-->
<%
' Here is the structure
' ----------------------------------------
'  SELECT CASE Process
'    CASE "LIST"                            ' [2.0]
'     select case aAction
'     case "ADD NEW " & ucase(aNiceTableName): response_redirect 2.21, RootFileName & "?ID=0", ""
'     case "FIRST":   response_redirect 2.22, RootFileName & "?Page=1", ""
'     case "PREV" :   response_redirect 2.23, RootFileName & "?Page=" & session("nPage") - 1, ""
'     case "NEXT" :   response_redirect 2.24, RootFileName & "?Page=" & session("nPage") + 1, ""
'     case "LAST" :   response_redirect 2.25, RootFileName & "?Page=" & session("nPageCount"), ""
'     case "SHOW" :   response_redirect 2.26, RootFileName,  ""
'     case "REPORT":  response_redirect 2.27,"incgeneratereport.asp", ""
'     case "FILTER":  response_redirect 2.28, RootFileName & "?ID=-1", ""
'    CASE "FORM"                            ' [3.0]
'     select case aAction
'      case "EDIT", "NEW"                   ' [3.1]
'      case "SAVE CHANGES", "SAVE & CLOSE"  ' [4.0]
'       Insert New Record
'         Check4Duplicates Request          ' [4.1]
'         chkValidation    pRS              ' [4.2]
'         Check4Validation Request          ' [4.5]
'         INSERT ...                        ' [4.3]
'         Session("ID")  = pRS(strIDField)  ' [4.4]
'         onAfterUpdate    pRS(strIDField)  ' [4.8]
'       Update Old Record
'         chkValidation    pRS              ' [4.5]
'         Check4Validation Request          ' [4.5]
'         archiveIt        pRS              ' [4.6]
'         saveChanges      pRS, aSQL        ' [4.7]
'         onAfterUpdate    pRS(strIDField)  ' [4.8]
'      case "VALIDATE"                      ' [5.0]
'         Check4Validation Request          ' [4.5]
'      case "DELETE"                        ' [6.0]
'      case "REPORT"                        ' [7.0]
'      case "MARK COMPLETE"                 ' [3.8]
'      case "LIST"                          ' [2.9]
'      case "APPLY FILTER"                  ' [7.5]
'      case "CLEAR FILTER"                  ' [7.6]
'      case else
'    CASE "MARK COMPLETE"                   ' [8.0]
'     select case aAction
'      case "SHOW"                          ' [8.1]
'      case "MARK COMPLETE"                 ' [8.2]
'      case "CANCEL"                        ' [8.3]
'      case else
'    CASE "ANOTHER PROCESS"                 ' [9.0]
'
' debugging controls
' ----------------------------------------
'
'    nNoRedirect_No  =  4.8  '  onAfterUpdate
'    nNoRedirect_No  =  4.9  '  Save Old Rec

     bNoEmails       =  true
     bNoEmails       =  false

     bDebug          =  true
     bDebug          =  false

     session("RO")   = "NO"

' ==========================================================================================
' SET SQL TABLE NAME, ID FIELD, LIST/EDIT/SORT FIELD NAMES AND WHERE CLAUSE
' ==========================================================================================

'    usesMarkedComplete = true
     usesMarkedComplete = false

'    usesEmail = true
     usesEmail = false

     usesMenus = true
'    usesMenus = false

     strUpNav = "tables.asp"

     tablePrefix = "tbl"

 If (len(trim(request("t"))) > 0) then
     strTableName            = request("t")
     strWhere4List           = "1=1"
     Session("TableName")    = strTableName
     Session("Where4List")   = strWhere4List
     Session("ShowRecords")  = 10
   else
     strTableName     = Session("TableName")
     end if

     aNiceTableName   = mid(strTableName,len(tablePrefix)+1)

   Select Case lcase(strTableName)
     Case "tblusers":        strIDField = "UID"
     Case "tblagencies":     strIDField = "AID"
     Case "tblcommittees":   strIDField = "CID"
     Case "tblgroups":       strIDField = "GID"
     Case Else  :            strIDField = aNiceTableName & "ID"
       End Select

     strWhere4List    =  Session("Where4List")
     strDefaultSort   =  strIDField     '
     strDefDirection  = "ASC"          '

' Create Special Selection Rules Here
' -----------------------------------

     strFields4List   = "*"

'If (Session("PermissionLevel") = "CMS") then
'    strFields4List   = "ConsultationsID, ConsultNo, ConsultType, ReceivedDate, ConcurredDate, CharterDate, CharterTerminated, Comments, CMORemarks"
'Else
'    strFields4List   = "ConsultationsID, ConsultNo, ConsultType, ReceivedDate, ConcurredDate, CharterDate, CharterTerminated,           CMORemarks"
'End If
' --------------------------------------

     strFields4Add    = "*"
'    strFields4Add    = "CNO, ConsultType, ConsultNo, ReceivedDate, ConcurredDate, CharterDate, CharterTerminated,           CMORemarks"

     strFields4Edit   =  strFields4List
     strFields4Filter =  strFields4Add

    If useMarkedComplete then
     strFields4Sys    = "ChangedAt, ChangedBy, CreatedAt, CreatedBy, MarkedCompleteAt, MarkedCompleteBy"
      else
     strFields4Sys    = "ChangedAt, ChangedBy, CreatedAt, CreatedBy"
       end if

     strFieldsNot4List= "CreatedAt, CreatedBy"
' -------------------------------------------------------------------------------------------

  Function SetCustomFields( aName, aValue )
       Dim pRS, aStr:    aStr = ""
    select case aName

     case "ConsultType": aStr = fmtSelect( aFill, aName, aValue)
       end select
           SetCustomFields = aStr
  End Function

' -------------------------------------------------------------------------------------------
  Function SetRequiredHiddenFields( pFld )
       Dim pRS, aStr:    aStr = ""
    select case pFld.Name

      case "CNO":        aStr = "<INPUT type='Hidden' Name='CNO' Value='" & cnum(session("CNO")) & "'>"

      case "ConsultNo":' Get Max ConsultNo

                         aSQL = "SELECT Max(ConsultNo) AS MaxNo FROM tblConsultations WHERE CNO = " & session("CNO")
                     Set pRS  =  Conn.Execute(aSQL)
                         aStr = "<INPUT type='Hidden' Name='ConsultNo' Value='" & pRS("MaxNo") + 1 & "'>"
                         pRS.close
                     Set pRS  = Nothing
       end select
           SetRequiredHiddenFields = aStr
  End Function

' -------------------------------------------------------------------------------------------
  Function Check4Duplicates( Request )
       Dim pRS
'          aSQL = "SELECT * FROM tPerson  WHERE Logon = '" & replace(request("Logon"),"'", "''") & "'"
'      Set pRS  = conn.execute(aSQL)
'      if (pRS.EOF = false) then setValidation "Logon", "Duplicate Email Address", "* The email Address you entered in already in the system. Please enter another."

  End Function

' -------------------------------------------------------------------------------------------
  Function Check4Validation( pRS)

'          chkRequired "Phone",        "Phone Number",  "* Please enter a Phone Number for this user."

        if (("" = pRS("CharterTerminated") & "") and trim(lcase(pRS("ConsultType"))) = "termination") then
            setValidation "CharterTerminated","CharteredTerminated when ConsultType is Termination", "can not be empty"
            end if

  End Function

' -------------------------------------------------------------------------------------------

    if ("" = strFields4Add  & "") then strFields4Add  = "*"
    if ("" = strFields4Edit & "") then strFields4Edit = "*"

    if (bDebug = true) then
        response.write "strTable Variables                    <br>"
        response.write "--------------------------------------<br>"
        response.write "TablePrefix: "    & TablePrefix    & "<br>"
        response.write "strTableName: "   & strTableName   & "<br>"
        response.write "aNiceTableName: " & aNiceTableName & "<br>"
        response.write "strIDField: "     & strIDField     & "<br>"
        response.write "strFields4Sys: "  & strFields4Sys  & "<br>"
        response.write "strFields4List: " & strFields4List & "<br>"
        response.write "strFields4Edit: " & strFields4Edit & "<br>"
        response.write "strFields4Add: "  & strFields4Add  & "<br>"
        response.write "strWhere4List: "  & strWhere4List  & "<br>"
        response.write "strFields4Sort: " & strFields4Sort & "<br>"
        response.write "usesMarkedComplete: " & usesMarkedComplete & "<br>"
        response.write "usesEmail: " & usesEMail & "<br>"
        response.write "usesMenus: " & usesMenus & "<br>"
        end if

' ==========================================================================================
' INITIALIZATION                          ' [1.0]
' ==========================================================================================

' Control Settings
' ------------------------------------------------------------------------

    RootFileName =  Request.ServerVariables("PATH_INFO")

if (len(trim(session("RO"))) = 0) then
    session("RO")= "YES"
    end if

    bRO          =  ucase(left(Session("RO"),1)) = "Y"
If (bRO) then
    aBtnOff      = " style=""color:#888888"" onclick=""return false"""
    end if

    PageTitle    =  aNiceTableName & " Information"
    BaseColor    = "DarkBlue" ' DarkBlue Maroon  DarkGreen
    BorderColor  = "#C0C0C0"

    aFill        =  vbCrLf & "            "
    aHTML        =          "<td align=""top"" ><font size=""3"">{FieldLabel}:&nbsp;</font></td>" & aFill & "  "
    aHTML        = aHTML  & "<td align=""top"" ><font size=""3"">{FieldValue}</font></td>"

If (len(trim(Session("ShowRecords"))) = 0) then Session("ShowRecords") = 10
If (len(trim(Request("ShowRecords"))) > 0) then
    Session("ShowRecords") = left(Request("ShowRecords"), instr(Request("ShowRecords") & ",", ",") - 1)
    If int(Session("ShowRecords")) > 500 then Session("ShowRecords") = 100
End If

' Set up for processing
' ------------------------------------------------------------------------

        NewProcess = Session("Process")
    if (Request.QueryString("Process") > "") then NewProcess = Request.QueryString("Process")

If (len(trim(request.querystring("ID"))) > 0) then
        Session("ID")      = cnum(request.querystring("ID"))
    if (NewProcess > "") then
        Process = NewProcess
      else
        Process = "FORM"
        end if
  else
        Session("ID")      = 0
    if (session("Process") > "" and NewProcess <> "FORM") then
        Process = NewProcess
      else
        Process = "LIST"
        end if
    end if
        Session("Process") = null

        aAction = ucase(trim(request("btn")))

        bFirstTime = (strTableName <> session("table") & "" )
        bFirstTime =  bFirstTime or (0 = instr( 1, strFields4List, session("SortType"), 1))
    if (bFirstTime) then
        session("table")         = strTablename
        session("SortType")      = ""
        session("SortDirection") = ""
        end if

' ==========================================================================================
' BEGIN HTML HEADER FOR EVERY PAGE
' ==========================================================================================
%>
<!doctype HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<meta NAME="Author" CONTENT="80/20 Data Co.">
<meta NAME="Generator" CONTENT="Microsoft FrontPage 6.0">
<meta name="ProgId" CONTENT="FrontPage.Editor.Document">
<title><%= PageTitle %></title>
</head>

<!-- ---------------------------------------------------------------------------------------- -->

<!--#INCLUDE FILE="incheader.asp"-->
<body>
<!--#INCLUDE FILE="incbodyline.asp"-->
<% Up = strUpNav %>
<!--#INCLUDE FILE="incnav.asp"-->
<BR>

<!-- ---------------------------------------------------------------------------------------- -->
<% ' Put Code for top page activity here. %>

<table bordercolor="<%= BaseColor %>" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td>
    <table>
      <tr>
        <td>
        <%

' Validation and System Messages
' ------------------------------------------------------------------------

      bLstErr = (Session("valCount") > 0 AND Session("valMessage") > "")
  if (bLstErr) then
%>
        <table width="1" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="images/spacer.gif"></td>
            <td><hr align="left" color="red" width="400" size="1"></td>
          </tr>
          <tr>
            <td width="15"></td>
            <td><font color="red" size="3"><b>Violated Validation Rules: </b>
            </font></td>
          </tr>
          <tr>
            <td width="80"></td>
            <td><font color="red" size="3"><b><%= Session("valMessage") %></b></font></td>
          </tr>
          <tr>
            <td><img src="images/spacer.gif"></td>
            <td><hr align="left" color="red" width="400" size="1"></td>
          </tr>
        </table>
        <%
    Else
      bMsg    = (5 < len(Session("Where4List"))) and (Process = "LIST" or (Process = "FORM" and Session("ID") = -1))
  if (bMsg) then
%>
        <table width="1" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="images/spacer.gif"></td>
            <td><hr align="left" color="green" width="400" size="1"></td>
          </tr>
          <tr>
            <td width="80"></td>
            <td><font color="green" size="3"><b>Filter is On: <%= Session("Where4List") %></b></font></td>
          </tr>
          <tr>
            <td><img src="images/spacer.gif"></td>
            <td><hr align="left" color="green" width="400" size="1"></td>
          </tr>
        </table>
        <%
      End If

      bMsg    = (Session("Message") > "")
  if (bMsg) then
%>
        <table width="1" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="images/spacer.gif"></td>
            <td><hr align="left" color="blue" width="400" size="1"></td>
          </tr>
          <tr>
            <td width="80"></td>
            <td><font color="blue" size="3"><b><%= Session("Message") %></b></font></td>
          </tr>
          <tr>
            <td><img src="images/spacer.gif"></td>
            <td><hr align="left" color="blue" width="400" size="1"></td>
          </tr>
        </table>
        <%
      End If
      End If
      Session("Message") = ""

' Do Processes
' ------------------------------------------------------------------------

   if (bDebug) then response.write "[1.0] Process is " & ucase(Process) & "<br>"

SELECT CASE Process

' ==========================================================================================
' LIST PROCESSING              [2.0]
' ==========================================================================================

  CASE "LIST"                             ' [2.0]

' Button Action
' ------------------------------------------------------------------------

    clrValidation 0

if (bDebug) then
    response.write "[2.0] PROCESS: LIST"
    response.write ", Btn:  '"   & request("btn") & "'"
    response.write ", Action: '" & aAction & "'"
    response.write ", Page: "    & CLng(Request.QueryString("Page"))
    response.write ", Sort: "    & request("Sort") & "<br>"
    end if

IF (len(aAction) > 0) then

    select case aAction
     case "ADD NEW " & ucase(aNiceTableName)
                    response_redirect 2.21, RootFileName & "?ID=0", ""
     case "FIRST":  response_redirect 2.22, RootFileName & "?Page=1", ""
     case "PREV" :  response_redirect 2.23, RootFileName & "?Page=" & session("nPage") - 1, ""
     case "NEXT" :  response_redirect 2.24, RootFileName & "?Page=" & session("nPage") + 1, ""
     case "LAST" :  response_redirect 2.25, RootFileName & "?Page=" & session("nPageCount"), ""
     case "SHOW" :  response_redirect 2.26, RootFileName,  ""
     case "FILTER": response_redirect 2.27, RootFileName & "?ID=-1", ""
     case "REPORT": response_redirect 2.28,"incgeneratereport.asp", ""
    end select
End if

' ------------------------------------------------------------------------
' Get the current page
' ------------------------------------------------------------------------

' SQL SELECT
' ------------------------------------------------------------------------

  if ("" = strWhere4List  & "") then strWhere4List =  "1 = 1"
  if ("" = strFields4List & "") then strFields4List = "*"
  if ("" = strFields4Sort & "") then strFields4Sort = strIDField

'            aSQL =    "SELECT "  &              strFields4List     '*.(40426.01.1
             aSQL =    "SELECT "  & fmtSQL_Flds( strFields4List)    ' .(40426.01.1
'     aSQL = aSQL &     " FROM "  & strTableName                    '*.(40426.01.2
      aSQL = aSQL &     " FROM [" & strTableName & "]"              ' .(40426.01.2
      aSQL = aSQL &    " WHERE "  & strWhere4List

' SQL SELECT Customization
' ------------------------------------------------------------------------

' If (Session("PermissionLevel") = "CMS") then
'            aSQL =    "SELECT ConsultationsID, CNo, ConsultType, ConsultNo, ReceivedDate "
'     aSQL = aSQL &         ", ConcurredDate, CharterDate, CharterTerminated, Comments "
'     aSQL = aSQL &         ", CMORemarks, CreatedAt, CreatedBy ,ChangedAt, ChangedBy "
'     aSQL = aSQL &     " FROM tblConsultations "
'     aSQL = aSQL &    " WHERE CNO = " & cnum(session("CNo"))
'  ** aSQL = aSQL & " ORDER BY ConsultNo DESC "
'   Else
'            aSQL =    "SELECT ConsultationsID, CNo, ConsultType, ConsultNo, ReceivedDate "
'     aSQL = aSQL &         ", ConcurredDate, CharterDate, CharterTerminated, "
'     aSQL = aSQL &         ", CMORemarks,  CreatedAt, CreatedBy ,ChangedAt, ChangedBy "
'     aSQL = aSQL &     " FROM tblConsultations "
'     aSQL = aSQL &    " WHERE CNO = " & cnum(session("CNo"))
' **  aSQL = aSQL & " ORDER BY ConsultNo DESC "
'     End If

' SQL ORDER BY
' ------------------------------------------------------------------------

If (len(trim(request("Sort"))) > 0) then

    If (session("SortType") = request("Sort")) then
         if (Session("SortDirection") = "ASC") then
             SortDirection = "DESC"
           else
             SortDirection = "ASC"
             end if
     End If
             SortType = request("Sort")
Else
    If (len(trim(Session("SortType"))) > 0) then
             SortType = trim(Session("SortType"))
           else
             SortType = ""
             end if
    If (len(trim(Session("SortDirection"))) > 0) then
             SortDirection = trim(Session("SortDirection"))
           else
             SortDirection = ""
             end if
End If

If ("" = SortType      & "") then SortType      =  strDefaultSort
If ("" = SortDirection & "") then SortDirection =  strDefDirection

If (bDebug = true) then
    response.write "[2.3] SortType: " & SortType & " Direction: " & SortDirection & "<BR>"
End if

    Session("SortDirection") = SortDirection
    Session("SortType")      = SortType

    aSQL = aSQL & " ORDER BY " & SortType & " " & SortDirection

' EXECUTE SQL
' ------------------------------------------------------------------------

    Session("SQLStr") = aSQL

If (bDebug = true) then
    response.write "[2.4] SQL: " & aSQL & "<BR>"
'   response.write "[2.4] CS: "  &  SQLConnectionStr & "<br>"
'   response.end
End If

Set pRS = Server.CreateObject("ADODB.Recordset")
    pRS.CursorLocation = 3 ' adUseClient

    pRS.Open CleanSQL(aSQL), SQLConnectionStr

If (pRS.EOF) Then

    pRS.Close
Set pRS = Nothing

'   setValidation "","[2.4] SQL: "   & aSQL, ""
'   setValidation "","ERROR: No records found.", ""

    Session("message") = "No Records Found."
    response_redirect 2.4, "incmessage.asp", ""
    response.end

End If

' PAGINATE RESULT
' ------------------------------------------------------------------------

     nRecCount = pRS.RecordCount

' Tell recordset to split records in the pages of our size

    pRS.PageSize = Session("ShowRecords")

' Set the page number to list

    nPage = CLng(Request.QueryString("Page"))
    Session("nPage") = nPage

' Get the number of pages

    nPageCount            = pRS.PageCount
    Session("nPageCount") = nPageCount

' Make sure that the Page parameter passed to us is within the range

If (nPage < 1 Or nPage > nPageCount) Then nPage = 1

%>
        <table border="1" bordercolor="<%=BorderColor%>" cellspacing="0" style="border-collapse: collapse" cellpadding="0">
          <form method="POST" action="<%= RootFileName %>">
            <!-- BEG OF TOP LIST BAR ------------------------------------------------------------------------------------- -->
    <% If usesMenus then %>
            <tr>
              <td bgColor="<%= BaseColor %>" align="left" valign="middle" colspan="20">
              <%

          If (ucase(left(Session("RO"),1)) <> "Y") then  %>
              <input type="submit" value="Add New <%= aNiceTableName %>" name="btn" <%'= aBtnOff %>><%
              End If %> <input type="submit" value="First" name="btn">
              <input type="submit" value="Prev" name="btn">
              <input type="submit" value="Next" name="btn">
              <input type="submit" value="Last" name="btn">&nbsp;
              <input type="submit" value="Report" name="btn">&nbsp;
              <input type="submit" value="Filter" name="btn">&nbsp;
              <input type="submit" value="Show" name="btn"> &nbsp;<input type="text" value="<%= session("ShowRecords") %>" size="3" name="ShowRecords">
              <font size="2" color="#FFFFFF">&nbsp;<%= "   " & nPage & " of " & nPageCount & " pages" %>
              </font></td>
            </tr>
    <% End If %>
            <!-- COLUMN HEADINGS ----------------------------------------------------------------------------------------- -->
            <%
                response.write aFill & "<TR VALIGN='top' ALIGN='left'>"

'               response.write aFill & "   <td align='left' valign='top' ><font size='2'><b><a href=" & RootFileName & "?Sort=" & pRS(strIDField) & ">" & pRS(strIDField) & "</a></b></font>&nbsp;</td>"
    For Each pFld in pRS.Fields
'       If (0 = instr(1, strFields4Sys & ",", pFld.Name & ",", 1)) then
'       If (0 = instr(1, strFields4Sys,       pFld.Name, 1)) then
        If ( not inFlds( strFieldsNot4List,   pFld.Name)) then
			If ucase(right(pfld.Name,2)) = "ID" then
                response.write aFill & "   <td align='left' valign='top' ><font size='2'><b> &nbsp; </b></font>&nbsp;</td>"
			Else
                response.write aFill & "   <td align='left' valign='top' ><font size='2'><b><a href=" & RootFileName & "?Sort=" & pFld.Name & ">" & pFld.Name & "</a></b></font>&nbsp;</td>"
			End If
        end If
        Next
                response.write aFill & "</TR>"   %>
            <!-- BEG OF LIST FIELDS -------------------------------------------------------------------------------------- -->
            <%
' Position recordset to the page we want to see
' ------------------------------------------------------------------------

    pRS.AbsolutePage = nPage

' Loop through records until it's a next page or End of Records
' ------------------------------------------------------------------------

     Do While Not (pRS.EOF OR pRS.AbsolutePage <> nPage)

        aIDfld = pRS.Fields(0).Name
        aIDnum = pRS.Fields(0).Value:
		If session("ROLE") <> "RO" Then
	        aIDlink = "Edit"   '   aIDlink = aIDnum
	        response.write aFill & "<TR>"
		    response.write aFill & "   <td valign=""top""><font size=""2""><b><a href=""" & RootFileName & "?ID=" & aIDnum & """>" & aIDlink & "</a></b></font>&nbsp;</td>"
		Else
	        aIDlink = "View"   '   aIDlink = aIDnum
	        response.write aFill & "<TR>"
		    response.write aFill & "   <td valign=""top""><font size=""2""><b><a href=""" & RootFileName & "?ID=" & aIDnum & "&RO=ON"">" & aIDlink & "</a></b></font>&nbsp;</td>"
		End If

    For Each pFld in pRS.Fields
'       If (0 = instr(1, aIDfld & ", " & strFields4Sys,     pFld.Name, 1)) then
'       If ( not inFlds( aIDfld & ", " & strFields4Sys,     pFld.Name)) then
        If ( not inFlds( aIDfld & ", " & strFieldsNot4List, pFld.Name)) then    '

'           if (bDebug = true) then response.write "<td>" & I & " - " & pFld.Name & "</td>"

                aFld =  pFld.Value & ""
            If (aFld = "") then aFld = "&nbsp;"
                response.write aFill & "  <td valign=""top""><font size=""2"">" & aFld & "</font>&nbsp;</td>"
            end if
        Next
                response.write aFill & "</TR>"
            pRS.movenext
        loop

            pRS.Close
        Set pRS = Nothing  %>
            <!-- END OF LIST FIELDS -------------------------------------------------------------------------------------- -->
            <tr vAlign="top">
              <td bgColor="#F5F5F5" height="4" colspan="3"></td>
            </tr>
            <!-- BEG OF BOTTOM LIST BAR ---------------------------------------------------------------------------------- -->
    <% If usesMenus then %>
            <tr>
              <td bgColor="<%=BaseColor%>" align="left" valign="middle" colspan="20">
              <%

           If (ucase(left(Session("RO"),1)) <> "Y") then  %>
              <input type="submit" value="Add New <%= aNiceTableName %>" name="btn" <%'= aBtnOff %>><%
              End If %> <input type="submit" value="First" name="btn">
              <input type="submit" value="Prev" name="btn">
              <input type="submit" value="Next" name="btn">
              <input type="submit" value="Last" name="btn">&nbsp;
              <input type="submit" value="Report" name="btn">&nbsp;
              <input type="submit" value="Filter" name="btn">&nbsp;
              <!--input type="submit" value="Show" name="btn"> &nbsp;<input type="text" value="<%= session("ShowRecords") %>" size="3" name="ShowRecords"-->
              <font size="2" color="#FFFFFF">&nbsp;<%= "   " & nPage & " of " & nPageCount & " pages" %>
              </font></td>
            </tr>
    <% End If %>
            <!-- END OF BOTTOM LIST BAR ---------------------------------------------------------------------------------- -->
          </form>
        </table>
        <%
' ==========================================================================================
' FORM PROCESSING
' ==========================================================================================

  CASE "FORM"                             ' [3.0]

if (aAction = "") then
    Select case Session("ID")
      Case -1:    aAction = "FILTER"
      Case  0:    aAction = "NEW"
      Case Else:  aAction = "EDIT"
        End Select
    End If

if (bDebug) then
    response.write "[3.0] PROCESS: FORM"
    response.write ", Btn:  '"     & request("btn") & "'"
    response.write ", Action: "    & aAction
    response.write ", ValCount:  " & cnum(Session("valCount")) & "<br>"
    end if

if (instr("EDIT,NEW", aAction) = 0) then clrValidation 0

' -----------------------------------------------------------------------------------------
    select case aAction

     case "EDIT", "NEW", "FILTER"         ' [3.1]

' --------------------------------------------------------------------------------
%>
        <table border="1" bordercolor="<%= BorderColor %>" cellspacing="0" style="border-collapse: collapse" cellpadding="0">

          <form method="POST" action="<%= RootFileName %>?ID=<%= session("ID") %>">

            <!-- BEG OF TOP FORM BAR ------------------------------------------------------------------------------------- -->

            <tr vAlign="top">
              <td bgColor="<%= BaseColor %>" align="center" valign="middle" colspan="3" height="28">
              <p align="left">
              <%
        if (aAction = "FILTER") then %>
              <input type="submit" value="Apply Filter"     name="btn" title="Apply Filter"     <%'= aBtnOff %>>
              <input type="submit" value="Clear Filter"     name="btn" title="Clear Filter"     <%'= aBtnOff %>>
    <%    else
          if (ucase(left(Session("RO"),1)) <> "Y")          then %>
              <input type="submit" value="Save Changes"     name="btn" title="Save Changes"     <%'= aBtnOff %>>
              <input type="submit" value="Save &amp; Close" name="btn" title="Save &amp; Close" <%'= aBtnOff %>>
         <!-- <input type="submit" value="Validate"         name="btn" title="Validate"         <%'= aBtnOff %>>  -->
              <%
              if (instr("ADM,POC", session("Role")) > 0)    then %>
              <input type="submit" value="Delete"           name="btn" title="Delete"           <%'= aBtnOff %>>
              <%
              if (aAction <> "NEW" AND usesMarkComplete)    then %>
              <input type="submit" value="Mark Complete"    name="btn" title="Mark Complete"    <%'= aBtnOff %>>
              <%
                  end if  '-- NEW
              end if  '-- ADM or POC
          end if   '-- RO                                        %>
        <!--  <input type="submit" value="Report"           name="btn" title="Report">  -->
              <input type="submit" value="List"             name="btn" title="List">&nbsp;&nbsp;
        <% end if %>
              </td>
            </tr>
            <!-- END OF TOP FORM BAR ------------------------------------------------------------------------------------- -->

            <tr vAlign="top">
              <td bgColor="#F5F5F5" height="4" width="1%" colspan="3"></td>
            </tr>

            <!-- BEG OF FORM FIELDS -------------------------------------------------------------------------------------- -->
            <%

' Display Fields for One Record
' ------------------------------------------------------------------------

    Select Case aAction
      Case "NEW":    aSQL = "SELECT " & fmtSQL_Flds( strFields4Add   ) & " FROM [" & strTableName & "]"  ' .(40426.01.3
      Case "EDIT":   aSQL = "SELECT " & fmtSQL_Flds( strFields4Edit  ) & " FROM [" & strTableName & "]"  ' .(40426.01.4
      Case "FILTER": aSQL = "SELECT " & fmtSQL_Flds( strFields4Filter) & " FROM [" & strTableName & "]"  ' .(40426.01.5
        end Select

              aSQL = aSQL & " WHERE " & strIDField & " = " & Session("ID")

      if (bDebug = true) then response.write "[3.2] SQL: " & aSQL & "<br>"

      Set pRS  = Conn.Execute(aSQL)

          bNewRec = (pRS.EOF or 0 = Session("ID"))
      if (bNewRec) then
      Set pRS = getMTRecordset(aSQL)
          end if

      For Each pFld in pRS.Fields

'         If (0 = instr(1, strIDField & ", " &                                            strFields4Sys,  pFld.Name, 1 )) then
'         If ( not inFlds( strIDField & ", " &                                            strFields4Sys , pFld.Name)) then
          If ( not inFlds( strIDField & ", " & iif(aAction = "FILTER", strFieldsNot4List, strFields4Sys), pFld.Name)) then
'         -------------------------------------------------------------------------

                response.write aFill & "<TR>"

                              aVal = pFld.Value
            if (bNewRec) then aVal = ""
            if (bLstErr) then aVal = resetLastValue( pFld.Name, aVal)

                bAddFld = not (pFld.Attributes and &H00000040)                             ' NOT NULL, ie. Required
                bAddFld = bAddFld OR 0 < instr( 1, strFields4Add, pFld.name, 1)            ' OR in Fields4Add List

'           if (bAddFld AND aAction = "NEW") then                                          ' should it be this
'               if (aAction = "NEW" and 0 = instr( 1, strFields4Add, pFld.name, 1)) then   ' it was this

'           Fields for New Record being Added / Inserted
'           -----------------------------------------------
            if (bAddFld AND aAction = "NEW") then                                          ' but what about custom fields

                        aStr =                SetCustomFields( pFld.Name, "")
                if (    aStr = ""          )  then
                        aStr = aFill & "  " & SetRequiredHiddenFields( pFld)
                    if (aStr = aFill & "  ")  then
                        aStr = aFill & "  " & displayField(pFld, aVal, aHTML)
                        end if
                    end if

                response.write aStr

'           Fields for Old Record being Edited / Updated
'           -----------------------------------------------
            else

                response.write aFill & "<TR>"

                aStr = SetCustomFields( pFld.Name, aVal)
            if (aStr > "") then

                response.write aStr
              else

              select  case pFld.Name

                case "StartDate"
                response.write aFill & "  <td>StartDate: &nbsp</td>"
                response.write aFill & "  <td>&nbsp" & htmDate( pFld, aVal,2) & "</td>"

                case "EndDate"
                response.write aFill & "  <td>EndDate: &nbsp</td>"
                response.write aFill & "  <td>&nbsp" & htmDate( pFld, aVal,2) & "</td>"

                case "Custom"
%><!--
          <td bgcolor="maroon" align="right"><font color="white"><b>EndDate:&nbsp;</b></font></td>
          <td><%= htmDate(pFld, aVal,2) %>&nbsp;</td>
--><%
                case  else

                response.write aFill & "  " & displayField( pFld, aVal, aHTML)

                end select
                end if

            end if   ' ------------------------------------

                response.write         msgIfError( pFld.Name )                           ' Err Msg for Field

                response.write aFill & "</TR>"

        End If
'       -------------------------------------------------------------------------
    Next

        pRS.Close
    Set pRS = Nothing
 %>
            <!-- END OF FORM FIELDS -------------------------------------------------------------------------------------- -->

            <tr vAlign="top">
              <td bgColor="#F5F5F5" height="4" colspan="3"></td>
            </tr>

            <!-- BEG OF BOTTOM FORM BAR ---------------------------------------------------------------------------------- -->

            <tr vAlign="top">
              <td bgColor="<%= BaseColor %>" align="center" valign="middle" colspan="3" height="28">
              <p align="left">
              <%
        if aAction = "FILTER" then %>
              <input type="submit" value="Apply Filter"      name="btn" title="Apply Filter" <%'= aBtnOff %>>
              <input type="submit" value="Clear Filter"      name="btn" title="Clear Filter" <%'= aBtnOff %>>
    <%    else
          if (ucase(left(Session("RO"),1)) <> "Y")           then %>
              <input type="submit" value="Save Changes"      name="btn" title="Save Changes" <%'= aBtnOff %>>
              <input type="submit" value="Save &amp; Close"  name="btn" title="Save &amp; Close" <%'= aBtnOff %>>
       <!--   <input type="submit" value="Validate"          name="btn" title="Validate"     <%'= aBtnOff %>>  -->
              <%
              if (instr("ADM,POC", session("Role")) > 0)     then %>
              <input type="submit" value="Delete" name="btn" title="Delete" <%'= aBtnOff %>>
              <%
              if (aAction <> "NEW" AND usesMarkedComplete)   then %>
              <input type="submit" value="Mark Complete"     name="btn" title="Mark Complete" <%'= aBtnOff %>>
              <%
                  end if  '-- NEW
              end if  '-- ADM or POC
          end if   '-- RO                                         %>
        <!--  <input type="submit" value="Report"            name="btn" title="Report">  -->
              <input type="submit" value="List"              name="btn" title="List">&nbsp;&nbsp;
        <% end if %>
              </td>
            </tr>

            <!-- END OF BOTTOM FORM BAR ---------------------------------------------------------------------------------- -->
          </form>
        </table>
        <%
        clrValidation 0    ' After Form is displayed for NEW and EDIT

' -----------------------------------------------------------------------------------------

     case "SAVE CHANGES", "SAVE & CLOSE"  ' [4.0]

    if (session("ID") = 0) then  ' Insert New Record

' ---------------------------------------------------------------------------------

' Validate Input for Inserts only
' ------------------------------------------------------------------------

        Check4Duplicates Request           '[4.1]

'       aSQL = "SELECT TOP 1 " &              strFields4Add  & ", CreatedBy, CreatedAt FROM "  & strTableName        '*.(40426.01.7
        aSQL = "SELECT TOP 1 " & fmtSQL_Flds( strFields4Add) & ", CreatedBy, CreatedAt FROM [" & strTableName & "]"  ' .(40426.01.7

    if (bDebug = true) then response.write "[4.1] SQL: " & aSQL & "<br>"

    Set pRS  = Conn.Execute(aSQL)

        chkValidation pRS                 ' [4.2] Checks for NOT NULL, Dates and Numbers

        Check4Validation Request          ' [4.5] Table Specific Checks

    if (bDebug = true) then response.write "[4.2] Validation Errors: " & session("valCount") & "<br>"

    if (session("valCount") > 0) then
        response_redirect  4.2, RootFileName & "?ID=" & session("ID"), ""
        end if

' Insert New Record
' ------------------------------------------------------------------------
        dNow = now()

    if (strFields4Add = "*") then
        strFields4Add = getAllFieldsBut(pRS, strIDField & "," & strFields4Sys & ",CreatedBy,CreatedAt")
        end if

'                  aSQL = " INSERT "  & strTableName &  " (" &              strFields4Add  & ", CreatedBy, CreatedAt)"   '*.(40426.01.8
                   aSQL = " INSERT [" & strTableName & "] (" & fmtSQL_Flds( strFields4Add) & ", CreatedBy, CreatedAt)"   ' .(40426.01.8
            aSQL = aSQL & " VALUES (":  c = "  "
                   aFDs = "": k = 0

    For Each pFld in pRS.Fields
'       if (0 < instr(1, strFields4Add, pFld.Name, 1)) then
        if (     inFlds( strFields4Add, pFld.Name)) then
'           k = k+1:  response.write k & ") " & pFld.Name & "<br>"
'           aFDs = aFDs &        c   & fmtFValue( pRS(pFld.Name  ), request( pFld.Name ))
            aSQL = aSQL &        c   & fmtFValue( pRS(pFld.Name  ), request( pFld.Name )): c = ", "
        end if
        Next
            aSQL = aSQL &        c   & fmtFValue( pRS("CreatedBy"), Session("ChangedBy"))
            aSQL = aSQL &       ", " & fmtFValue( pRS("CreatedAt"), dNow                )
            aSQL = aSQL & "            )"

    if (bDebug = true) then
        response.write "<br>ADD: " & fmtSQL_Flds( strFields4Add) & ", CreatedBy, CreatedAt)" & "<br><br>"
        response.write "FLD: " & aFDs & "<br><br>"

        response.write "[4.3] SQL: " & aSQL & "<br>"
        response.end
        end if

     on error resume next: if (cnum(bDeBug) = true) then on error goto 0

        aSQL = CleanSQL( aSQL)
           Conn.Execute( aSQL)

    if (Err.Number <> 0) then

        response.write "<font color=red>Error SQL: " & aSQL & "</font><br>"

        setValidation "","[4.3] SQL: " & aSQL, ""
        setValidation "","ERROR: " & Err.Description, ""
        response_redirect  4.3, RootFileName & "?ID=0", ""
        end if

    set pRS  =  Server.CreateObject( "ADODB.Recordset" )
'       pRS.Open "SELECT " & strIDField & " FROM "  & strTableName &  " WHERE CreatedAt = '" & dNow & "'", Conn      '*.(40426.01.9
        pRS.Open "SELECT " & strIDField & " FROM [" & strTableName & "] WHERE CreatedAt = '" & dNow & "'", Conn      ' .(40426.01.9

    if (pRS.EOF) then
        setValidation "", "Record not Added. Please try again.", ""
        response_redirect  4.4, RootFileName & "?ID=" & session("ID"), ""
      else
        Session("Message")  = "Record Saved"
        end if

        Session("ID") = pRS(strIDField)   ' [4.4]
        Session("Message") = "Record ID " & Session("ID") & " Added."

        onAfterUpdate pRS( strIDField )   ' [4.8]

    if (Err.Number <> 0) then
        setValidation "","[4.8] Table Function: onAfterUpdate( " & pRS( strIDField) & " )", ""
        setValidation "","ERROR: " & Err.Description, ""
        response_redirect  4.8, RootFileName & "?ID=" & session("ID"), ""
        end if

     on error goto 0

' -----------------------------------------------------------------------------------------
      else   ' Save Old Record
' -----------------------------------------------------------------------------------------

    set pRS  =  Server.CreateObject( "ADODB.Recordset" )
'       pRS.Open "SELECT " &              strFields4Edit  & " FROM "  & strTablename &  " WHERE " & strIDField & " = " & Session("ID"), Conn  '*.(40426.02.1
        pRS.Open "SELECT " & fmtSQL_Flds( strFields4Edit) & " FROM [" & strTablename & "] WHERE " & strIDField & " = " & Session("ID"), Conn  ' .(40426.02.1

' Validate Input for Update only
' ------------------------------------------------------------------------

        chkValidation pRS                 ' [4.5] Checks for NOT NULL, Dates and Numbers

        Check4Validation Request          ' [4.5] Table Specific Checks

    if (session("valCount") > 0) then
        response_redirect 4.5, RootFileName & "?ID=" & session("ID"), ""
        end if

' Save Old Record
' ------------------------------------------------------------------------
     on error resume next: if (cnum(bDeBug) = true) then on error goto 0

'       archiveIt   pRS                   ' [4.6]

    if (Err.Number <> 0) then
        setValidation "","[4.6] SQL: "   & aSQL, ""
        setValidation "","ERROR: " & Err.Description, ""
        response_redirect  4.6, RootFileName & "?ID=" & session("ID"), ""
        end if

        saveChanges pRS, aSQL             ' [4.7]

    if (Err.Number <> 0) then
        setValidation "","[4.7] SQL: "   & aSQL, ""
        setValidation "","ERROR: " & Err.Description, ""
        response_redirect  4.7, RootFileName & "?ID=" & session("ID"), ""
        end if

        Session("Message") = "Record ID " & Session("ID") & " Saved."

        onAfterUpdate pRS( strIDField)    ' [4.8]

   if (Err.Number <> 0) then
        setValidation "","[4.8] Table Function: onAfterUpdate( " & RS( strIDField) & " )", ""
        setValidation "","ERROR: " & Err.Description, ""
        response_redirect  4.8, RootFileName & "?ID=" & session("ID"), ""
        end if

     on error goto 0

    end if

        pRS.Close
    set pRS = Nothing

' Return to FORM or LIST
' ------------------------------------------------------------------------

    if (aAction = "SAVE & CLOSE" and session("valCount") = 0) then

        response_redirect  4.91, RootFileName, ""
      else

        response_redirect  4.92, RootFileName & "?ID=" & Session("ID"), "FORM"
        end if

' -----------------------------------------------------------------------------------------

     case "VALIDATE"                      ' [5.0]

    set pRS  =  Server.CreateObject( "ADODB.Recordset" )
'       pRS.Open "SELECT " &              strFields4Edit  & " FROM "  & strTablename &  " WHERE " & strIDField & " = " & Session("ID"), Conn  '*.(40426.02.2
        pRS.Open "SELECT " & fmtSQL_Flds( strFields4Edit) & " FROM [" & strTablename & "] WHERE " & strIDField & " = " & Session("ID"), Conn  ' .(40426.02.2

        chkValidation pRS

        Check4Validation pRS

    if (session("valCount") = 0) then
        Session("Message")  = "All Fields in Record ID " & Session("ID") & " Are Valid."
        end if

        response_redirect 5, RootFileName & "?ID=" & Session("ID"), ""

' -----------------------------------------------------------------------------------------

     case "DELETE"                        ' [6.0]

'       aSQL = "DELETE "  & strTablename &  " WHERE " & strIDField & " = " & Session("ID")    '*.(40426.02.3
        aSQL = "DELETE [" & strTablename & "] WHERE " & strIDField & " = " & Session("ID")    ' .(40426.02.3

    if (bDebug = true) then
        response.write "[6.0] SQL: " & aSQL & "<br>"
      else
        Conn.Execute(aSQL)
        end if

        Session("Message") = "Record ID " & Session("ID") & " Deleted. " & iif(bDebug, "<font color=red>(Not if debugging)</font>", "")

        response_redirect 6.0, RootFileName, ""

' -----------------------------------------------------------------------------------------

     case "REPORT"                        ' [7.0]

        response_redirect 7.0, "incgeneratereport.asp", ""

' -----------------------------------------------------------------------------------------

     case "MARK COMPLETE"                 ' [3.8]

        Session("Process") = "MARK COMPLETE"
        response_redirect 3.8, RootFileName & "?ID=" & session("ID"), ""

' -----------------------------------------------------------------------------------------

     case "LIST"                          ' [2.9]

        response_redirect 2.9, RootFileName, ""

' -----------------------------------------------------------------------------------------

     case "APPLY FILTER"                  ' [7.5]

    set pRS  =   Server.CreateObject( "ADODB.Recordset" )
'       pRS.Open "SELECT TOP 1 " &              strFields4Filter  & " FROM "  & strTablename &  " WHERE " & strIDField & " = " & Session("ID"), Conn  '*.(40426.02.4
        pRS.Open "SELECT TOP 1 " & fmtSQL_Flds( strFields4Filter) & " FROM [" & strTablename & "] WHERE " & strIDField & " = " & Session("ID"), Conn  ' .(40426.02.4

'       bldFilter pRS
        Session("Where4List") = bldFilter( pRS )
'       Session("Message") = "Filter is: " & Session("Where4List")

        pRS.Close
    set pRS = Nothing

        response_redirect 5.0, RootFileName, ""

' -----------------------------------------------------------------------------------------

     case "CLEAR FILTER"                  ' [7.6]

        Session("Where4List") = "1 = 1"

    if (bDebug = true) then
        response.write "[9.1] Clear Filter: " & "<br>"
      else

        end if

        response_redirect 6.0, RootFileName, ""

' -----------------------------------------------------------------------------------------

     case else

        Session("Message") = "Form Action System Error"
        response_redirect, 7.9. RootFileName & "?ID=" & session("ID"), ""

      end select

' ==========================================================================================
' MARK COMPLETE
' ==========================================================================================

  CASE "MARK COMPLETE"                    ' [8.0]

if (aAction = "") then aAction = "SHOW"

if (bDebug) then
    response.write "[8.0] PROCESS: MARK COMPLETE"
    response.write ", Btn:  '"  & Request.Form("btn") & "'"
    response.write ", Action: " & aAction
    end if

    clrValidation 0 'pRS

    select case aAction

' -----------------------------------------------------------------------------------------

     case "SHOW"                          ' [8.1]
%>
        <table border="1" bordercolor="<%= BorderColor %>" cellspacing="0" style="border-collapse: collapse" cellpadding="0" width="100%">
          <form method="POST" action="<%= RootFileName %>?ID=<%= session("ID") %>&Process=MARK COMPLETE">
            <tr vAlign="top">
              <td bgColor="#000080" align="center" valign="middle" colspan="3" height="28">
              <p align="left">
              <%
      if (instr("ADM,POC",session("Role")) > 0 AND usesMarkedComplete)      then %>
              <input type="submit" value="Mark Complete" name="btn" <%= aBtnOff %>>
              <%
          end if %> <input type="submit" value="Cancel" name="btn"> &nbsp;</td>
            </tr>
            <tr>
              <td>
              <div align="center">
                <center>
                <table border="3" cellpadding="3" cellspacing="0" bordercolor="#000080">
                  <tr>
                    <td bgColor="#000080"><font color="#FFFFFF"><font size="5">
                    <b>Congratulations.&nbsp; Your Data has Passed all
                    Validation Rules.</b></font> </font></td>
                  </tr>
                  <tr>
                    <td>
                    <p align="center"><b><br>
                    <% If usesMarkedComplete AND instr("POC",Session("Grouplist")) > 0 OR instr("ADM",Session("Grouplist")) > 0 OR Session("PersonLastName") = "Schinner" Then %>
                    </p>
                    <div align="center">
                      <table border="0" cellpadding="5" cellspacing="0" height="164">
                        <tr>
                          <td align="center" height="38"><b>Now that you have
                          completed the data and are ready to Approve it.<br>
                          Press the Mark Complete button below.</b></td>
                        </tr>
                        <tr>
                          <td align="center" height="16"><hr></td>
                        </tr>
                        <tr>
                          <td align="center" height="27">
                          <input type="submit" value="Mark Complete" name="btn" title="Mark you data complete.  Can only be UNMARKED by Admin Personnel">
                          </td>
                        </tr>
                        <tr>
                          <td height="14"><hr></td>
                        </tr>
                        <tr>
                        </tr>
                      </table>
                    </div>
                    <% End If %> </b></td>
                  </tr>
                </table>
                </center>
              </div>
              </td>
            </tr>
          </form>
          <%
' -----------------------------------------------------------------------------------------

     case "MARK COMPLETE"                 ' [8.2]

    If usesMarkedComplete then
'       aSQL = "UPDATE "  & strTableName &  " Set MarkedCompleteBy = '"& Session("ChangedBy") &"', MarkedCompleteAt = '"& Session("Now") & "' WHERE " & strIDField & " = " & Session("ID")  '*.(40426.02.5
        aSQL = "UPDATE [" & strTableName & "] Set MarkedCompleteBy = '"& Session("ChangedBy") &"', MarkedCompleteAt = '"& Session("Now") & "' WHERE " & strIDField & " = " & Session("ID")  ' .(40426.02.5
        Conn.Execute(aSQL)


    If (len(trim(session("POCEMAIL"))) > 0) then

        Set Mail = Server.CreateObject("Persits.MailSender")
            Mail.Host = "smtp.8020data.com"     ' Specify a valid SMTP server
            Mail.From = Session("personlogon") ' Specify sender's address
            Mail.FromName = Session("ChangedBy") & " (Automated)"    ' Specify sender's name mailto = session("POCEMAIL")

         do while len(trim(Mailto)) > 0
        If (InStr(Mailto, ";")) Then
            Address = Left(Mailto, InStr(Mailto, ";") - 1)
           'response.write "Address: " & Address
            Mail.AddAddress Address
            Mailto = Mid(Mailto, InStr(Mailto, ";") + 1)
          Else
            Mail.AddAddress Mailto
            Mailto = ""
        End If
            Loop

'           Mail.AddAddress session("POCEMAIL") ' Name is optional
            Mail.Subject = Session("FY") & " " & Session("SystemAcronym") & " / " &  Session("ApplicationReference") & " is complete."
            Mail.Body = "This report has been marked it complete.  Thank you."

         On Error Resume Next
            Mail.Send

        Set Mail = Nothing

            If (Err <> 0) Then
                Session("Message") = "Error encountered: " & Err.Description
                response_redirect 8.21, RootFileName & "?ID=" & session("ID"), ""
            End if
        End If

        Session("Message") = "Record ID " & Session("ID") & " Marked Complete."
        response_redirect 8.22, RootFileName & "?ID=" & session("ID"), "FORM"

    End If
' -----------------------------------------------------------------------------------------

     case "CANCEL"                        ' [8.3]

        Session("Message") = "Mark Complete Canceled"
        response_redirect 8.3, RootFileName & "?ID=" & session("ID"), "FROM"

' -----------------------------------------------------------------------------------------

     case else                            ' [8.4]

        Session("Message") = "Mark Complete Action Error"
        response_redirect 8.4, RootFileName & "?ID=" & session("ID"), ""

        end select

' ==========================================================================================
' ANOTHER PROCESS
' ==========================================================================================

  CASE "ANOTHER PROCESS"                 ' [9.0]

        Session("Message") = "Another Process Done"
        response_redirect 9, RootFileName & "?ID=" & session("ID"), ""

' ==========================================================================================
  CASE ELSE
' ==========================================================================================

        Session("Message") = "System Process Error"
        response_redirect 10, RootFileName, ""

   END SELECT

' ==========================================================================================
' SAVE FUNCTIONS
' ==========================================================================================

   function archiveIt( pRS )

            aFields = ""
        for i = 0 to pRS.fields.count - 1
            aField = pRS.fields(i).name
            aFields = aFields & ", "  & aField        ' .(40426.02.6
'           aFields = aFields & ", [" & aField & "]"  '*.(40426.02.6
            next

'              aSQL = "INSERT INTO "  & strTableName &  " (" & mid(aFields, 2) & ") "                                                         '*.(40426.02.7
               aSQL = "INSERT INTO [" & strTableName & "] (" & mid(aFields, 2) & ") "                                                         ' .(40426.02.7
'       aSQL = aSQL & "SELECT " &               mid(aFields,2)  & " FROM "  & strTableName &  " WHERE " & strIDField & " = " & session("ID")  '*.(40426.02.8
        aSQL = aSQL & "SELECT " &  fmtSQL_Flds( mid(aFields,2)) & " FROM [" & strTableName & "] WHERE " & strIDField & " = " & session("ID")  ' .(40426.02.8

    if (bDebug = true) then
        response.write "[4.6] SQL: " & aSQL & "<br>"
'       response.end
        end if

        Conn.Execute( aSQL)

     end function

' --------------------------------------------------------

   function saveChanges( pRS, aSQL )

    for i = 0 to pRS.fields.count - 1

'       If (0 = instr(1, strIDField & ", " & strFields4Sys,                          pRS.Fields(i).Name, 1)) then
'       If ( not inFlds( strIDField & ", " & strFields4Sys,                          pRS.Fields(i).Name)) then
        If ( not inFlds( strIDField & ", " & strFields4Sys & ",ChangedBy,ChangedAt", pRS.Fields(i).Name)) then

'           aSQL = aSQL & ", " & fmtFldValue(            pRS.Fields(i) )   '*.(40426.02.9
            aSQL = aSQL & ", " & fmtSQL_FldEqRequestVal( pRS.Fields(i) )   ' .(40426.02.9
            end if
        next

            aSQL = aSQL & ", " & "ChangedAt = '"       &  now() & "'"
            aSQL = aSQL & ", " & "ChangedBy = '"       &  Session("ChangedBy") & "'"

        If (usesMarkedComplete) then
            aSQL = aSQL & ", " & "MarkedCompleteAt = " & "NULL"
            aSQL = aSQL & ", " & "MarkedCompleteBy = " & "NULL"
            End If

'           aSQL = "UPDATE "  & strTableName &  " SET " & mid(aSQL, 3) & " WHERE " & strIDField & " = " & session("ID")  '*.(40426.03.1
            aSQL = "UPDATE [" & strTableName & "] SET " & mid(aSQL, 3) & " WHERE " & strIDField & " = " & session("ID")  ' .(40426.03.1

    if (bDebug = true) then
        response.write "[4.7] SQL: " & aSQL & "<br>"
'       response.end
        end if

        aSQL = CleanSQL( aSQL)
           Conn.Execute( aSQL)

    end function

' ------------------------------------------------------------------------------------------ '

Function OnAfterUpdate( nID )   ' -------- .(40408.02.1 BEG  Added by Bruce  --------------- '
    Dim aSQL, pRS

'   Set Conn = Server.CreateObject("ADODB.Connection")
'       Conn.ConnectionTimeout = Session("Conn_ConnectionTimeout")
'       Conn.CommandTimeout    = Session("Conn_CommandTimeout")
'       Conn.Open                Session("Conn_ConnectionString"), Session("Conn_RuntimeUserName"), Session("Conn_RuntimePassword")

   'If (Termination) then update GeneralInfo fields and Mark Consultation Applied
   '-----------------------------------------------------------------------------------
    If (Session("ShowPrevFY") = "YES" and Session("FY") = Session("CurrentFY") and len(trim(Session("CMORolloverDate"))) = 0)  then

       'There is no General Info at this point.
       'So no updating now, will update during rollover
       'response.write "Same FY" & "<BR>" & Session("ShowPrevFY") & "<BR>"
       'response.write Session("FY") & "<BR>"
       'response.write Session("CurrentFY") & "<BR>"
       'response.write len(trim(Session("CMORolloverDate"))) & "<BR>"

    else

        Session("ConsultType")       = Request("ConsultType")
        Session("CharterDate")       = Request("CharterDate")
        Session("CharterTerminated") = Request("CharterTerminated")

    If (Session("ConsultType") <> "Termination" AND Session("CharterDate") <> "")  then

            If (Session("ConsultType") = "Admin Inactive" AND Session("CharterDate") <> "")  then
                aSQL = "UPDATE tblGeneralInfo Set CommitteeStatus = 'AdminInact' WHERE CID = " & Session("CID")

                Conn.Execute(CleanSQL(aSQL))
            Else
                aSQL = "UPDATE tblGeneralInfo Set CurrentCharterDate = '" & Session("CharterDate") & "', DateOfRenewalCharter = '" & DateAdd("yyyy",2,CDate(Session("CharterDate"))) & "', CommitteeStatus = 'Chartered' WHERE CID = " & Session("CID")

                Conn.Execute(CleanSQL(aSQL))
            End If
        End if

        If (Session("CharterTerminated") <> "") then
            If (Session("ConsultType") = "Termination") then
                aSQL = "UPDATE tblGeneralInfo Set ActualTerminationDate = '" & Session("CharterTerminated") & "', TerminatedThisFY = 'Yes', CommitteeStatus = 'Terminated' WHERE CID = " & Session("CID")
            else
                aSQL = "UPDATE tblGeneralInfo Set DateToTerminate = '" & Session("CharterTerminated") & "' WHERE CID = " & Session("CID")
            end if

        if (bDebug) then
            response.write "[4.8] SQL: " & aSQL & "<br>"
            end if

         on error resume next

            Conn.Execute(CleanSQL(aSQL))

        if (Err.Number <> 0) then
            setValidation "","[4.8] SQL: "   & aSQL, ""
            setValidation "","ERROR: " & Err.Description, ""
'           response_redirect  4.8, RootFileName & "?ID=" & session("ID"), ""
            end if

         on error goto 0

        End If
    end if

'===========================================
If (useEmail) then

   'Create Email

    aSQL     = "Select EMail FROM tblAgencies where AID= " & Session("AID") & " and Email is not null"
    Set pRS  =  Conn.Execute(aSQL)
    sTo      = ""
    If (not pRS.EOF AND Not IsNull(pRS("Email"))) then
        sTo  =  pRS("Email") + ";"
    End If
    Set pRS  =  Nothing

    sTo      =  sTo + "charles.howton@gsa.gov;" + "kennett.fussell@gsa.gov;" + "maggie.weber@gsa.gov;" + "btroutma@8020data.com;"
    sFrom    = "FACASupport@8020data.com" ' session("EmailFrom") ' eg kennett.fussell@gsa.gov
    sSubject = "New or Updated FACA Consultation for: " & Session("CNo") & "-" & Session("CommitteeName")
    sMessage =  ucase(session("logon")) & " has just updated the consultations for this committee to the FACA database at " & Now() & "."

if (cnum(bNoEmails) = true) then
    Session("Message") = Session("Message") & "<br>" & vbCrLf & "Emails NOT sent to " & sTo
    exit function
    end if

    If (session("emailtype") = "ASPEMAIL") then
        SendEMail sFrom, sTo,sSubject, sMessage, "consultations.asp"
    Else
        SendCDONTSMail  sFrom, sTo,sSubject, sMessage, "consultations.asp"
    End if

End If

End Function  ' --- .(40408.02.1 END  Added by Bruce  -------------------------------------- '

' ==========================================================================================
' BEGIN HTML FOOTER FOR EVERY PAGE
' ==========================================================================================
%> </td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    <p><br>
    </td>
  </tr>
</table>
<!-- ---------------------------------------------------------------------------------------- -->
<!--#INCLUDE FILE="incfooter.asp"-->

</body>

</html>
<%

Function response_redirect( n, aPage, aProcess)

' if (bDebug and         2.21        <> n) then
  if (bDebug and cnum(nNoRedirect_No) = n) then
      response.write "<font color=maroon>[" & n & "] Redirecting to " & aPage & aProcess & "</font><br>"
      response.end
      end if

  if (aProcess > "") then
      session("Process") = aProcess
      aProcess = "&Process=" & aProcess
      end if

      response.redirect aPage & aProcess

  End Function
%>