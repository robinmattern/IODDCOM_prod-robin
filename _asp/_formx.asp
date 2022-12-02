<% Response.Buffer = True %>

<!--#INCLUDE FILE="_incExpires.asp"    -->
<!--#INCL UDE FILE="_incLogonCheck.asp" -->
<!--#INCLUDE FILE="_inccommonfunctions.asp"-->
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<%

If (Session.TimeOut <> 60) then Session.TimeOut = 60

' Usage:
' For List:   Call RootFileName & ""
' For Form:   Call RootFileName & "?ID=xxxx"       where xxxx is the strTableName & "ID" value
' For Add:    Call RootFileName & "?ID=0           where 0 means add
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
'    generatereport.asp
'    generatefile.asp

' Here is the structure
' ----------------------------------------
'  SELECT CASE Process
'
'    CASE "LIST"                            ' [2.0]
'     select case aAction
'     case "ADD NEW RECORD" : response_redirect 2.21, RootFileName & "?ID=0", ""
'     case "FIRST":   response_redirect 2.22, RootFileName & "?Page=1", ""
'     case "PREV" :   response_redirect 2.23, RootFileName & "?Page=" & session("nPage") - 1, ""
'     case "NEXT" :   response_redirect 2.24, RootFileName & "?Page=" & session("nPage") + 1, ""
'     case "LAST" :   response_redirect 2.25, RootFileName & "?Page=" & session("nPageCount"), ""
'     case "SHOW" :   response_redirect 2.26, RootFileName,  ""
'     case "REPORT":  response_redirect 2.27,"generatereport.asp", ""
'     case "FILTER":  response_redirect 2.28, RootFileName & "?ID=-1", ""
'
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

     bNoEmails               =  cnum(bNoEmails)
'    bNoEmails               =  true
'    bNoEmails               =  false

     bDebug                  =  0 < instr(1, request.querystring(),"debug",1)
'     bDebug                  =  true
    bDebug                  =  false

' -------------------------------------------------------------------------------------------
                                                 session("RO") = "NO"
 if (""  =              session("RO") & "") then session("RO") = "NO"
 if (""  <  request.querystring("RO") & "") then session("RO") =  iif (request.querystring("RO") = "ON", "YES", "NO")
 if (         session("Role") = "RO"      ) then session("RO") = "YES"
 if (len(trim(session("RO"))) = 0         ) then session("RO") = "YES"

' -------------------------------------------------------------------------------------------

 If (len(trim(request("t"))) > 0) then
     strTableName                    = request("t")
   else
 if ("" = strTableName & "") then
     strTableName                    = Session("TableName") & ""
     end if
     end if

 if (strTableName = "") then                                            ' .(40621.01.1  BEG
     response.write "<font color=red><b>Error No Table Name!</b></font><br>"
     clrValidation 0

     Session("TableName")    = "tUNKNOWN"
     setValidation          "","ERROR: No Table name", ""
     response_redirect     4.3, RootFileName & "?ID=0", ""
     response.end
     end if                                                             ' .(40621.01.1  END

' Create Special Selection Rules Here
' -----------------------------------

 If (lcase(strTableName) = "tables") then       ' List Tables
     strTableName            = "SysObjects"
     aNiceTableName          = "Tables"
     strIDField              = "Name"
     Session("ShowRecords")  =  50

     strWhere4List           = "type = 'u' and substring(name,1,1) = 't'"
     strDefaultSort          =  strIDField                        '
     strDefDirection         = "ASC"                        '

     strFields4List          = "Name"
     usesMenus                = false
     End if

' ==========================================================================================
' SET SQL TABLE NAME, ID FIELD, LIST/EDIT/SORT FIELD NAMES AND WHERE CLAUSE
' ==========================================================================================

     usesMarkedComplete              =                                cnum(usesMarkedComplete) ' default is false
     usesEmail                       = iif("" = usesEmail & "", true, cnum(usesEmail))
     usesMenus                       = iif("" = usesMenus & "", true, cnum(usesMenus))
     strUpNav                        = trim(strUpNav & "")

 if ("" = tablePrefix       & "") then tablePrefix      = "t"
 if ("" = aNiceTableName    & "") then aNiceTableName   = mid(strTableName,len(tablePrefix)+1) ' .(40621.01.2
 if ("" = aTableName        & "") then aTableName       = mid(strTableName,len(tablePrefix)+1) ' .(40621.01.3
 if ("" = strIDField        & "") then strIDField       = aTableName & "ID"                    ' .(40621.01.4

 if ("" = strWhere4List     & "") then strWhere4List    =  Session("Where4List")
 if ("" = strDefaultSort    & "") then strDefaultSort   =  strIDField     '
 if ("" = strDefDirection   & "") then strDefDirection  = "ASC"          '

 if ("" = strFields4Add     & "") then strFields4Add    = "*"
 if ("" = strFields4Edit    & "") then strFields4Edit   = "*"
 if ("" = strFields4List    & "") then strFields4List   = "*"

 if ("" = strFields4Filter  & "") then strFields4Filter =  strFields4Add


' -------------------------------------------------------------

 If ("" = strFields4Sys     & "") then
 If (usesMarkedComplete) then
     strFields4Sys           = "ChangedAt, ChangedBy, CreatedAt, CreatedBy, MarkedCompleteAt, MarkedCompleteBy"
   else
     strFields4Sys           = "ChangedAt, ChangedBy, CreatedAt, CreatedBy"
     end if
     end if

'    strFieldsNot4List       = "CreatedAt, CreatedBy"

     Session("TableName")    =  strTableName
     Session("Where4List")   =  strWhere4List

' ------------------------------------------------------------------------------------------------------

    if (bDebug = true) then
        response.write "strTable Variables                    <br>"
        response.write "--------------------------------------<br>"
        response.write "TablePrefix: "        & TablePrefix        & "<br>"
        response.write "strTableName: "       & strTableName       & "<br>"
        response.write "aNiceTableName: "     & aNiceTableName     & "<br>"
        response.write "strIDField: "         & strIDField         & "<br>"
        response.write "strFields4Sys: "      & strFields4Sys      & "<br>"
        response.write "strFields4List: "     & strFields4List     & "<br>"
        response.write "strFields4Edit: "     & strFields4Edit     & "<br>"
        response.write "strFields4Add: "      & strFields4Add      & "<br>"
        response.write "strWhere4List: "      & strWhere4List      & "<br>"
        response.write "strFields4Sort: "     & strFields4Sort     & "<br>"
        response.write "usesMarkedComplete: " & usesMarkedComplete & "<br>"
        response.write "usesEmail: "          & usesEMail          & "<br>"
        response.write "usesMenus: "          & usesMenus          & "<br>"
        response.flush
        end if

' ==========================================================================================
' INITIALIZATION                          ' [1.0]
' ==========================================================================================

' Control Settings
' ------------------------------------------------------------------------

    RootFileName =  Request.ServerVariables("PATH_INFO")

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


    Session("ShowRecords") = left(Request("ShowRecords"), instr(Request("ShowRecords") & ",", ",") - 1)

'   Response.write "ReqShowRec: " & request("ShowRecords")

    If (len(trim(Session("ShowRecords"))) =   0) then Session("ShowRecords") =  10
    If (len(trim(Session("ShowRecords"))) >   0) then
    If (     int(Session("ShowRecords"))  > 500) then Session("ShowRecords") = 100
        End If

' Set up for processing
' ------------------------------------------------------------------------

        NewProcess         =  Session("Process")
    if (Request.QueryString( "Process") > "") then NewProcess = Request.QueryString("Process")

If (len(trim(request.querystring("ID"))) > 0) then
        Session("ID")      =  cnum(request.querystring("ID"))
    if (NewProcess > "") then
        Process            =  NewProcess
      else
        Process            = "FORM"
        end if
  else
        Session("ID")      =  0
    if (session("Process") > "" and NewProcess <> "FORM") then
        Process            =  NewProcess
      else
        Process            = "LIST"
        end if
    end if
        Session("Process") =  null

        aAction            =  ucase(trim(request("btn")))

        bFirstTime         = (strTableName <> session("table") & "" )
        bFirstTime         =  FirstTime or (0 = instr( 1, strFields4List, session("SortType"), 1))
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

<!--#INCLUDE FILE="_incheader.asp" ----------------------------------------------------------- -->

  <head>
    <meta HTTP-EQUIV="Content-Language" CONTENT="en-us">
    <meta HTTP-EQUIV="Content-Type"     CONTENT="text/html; charset=ISO-8859-1">
    <meta NAME="Author"                 CONTENT="80/20 Data Co.">
    <meta NAME="Generator"              CONTENT="Microsoft FrontPage 6.0">
    <meta NAME="ProgId"                 CONTENT="FrontPage.Editor.Document">

    <title><%= PageTitle %></title>

    <!--style>
    TR { font-family:Verdana; font-size=10pt; }
    TD { font-family:Verdana; font-size=10pt; }
    </style-->

  </head>
  <body>

<!--#INCLUDE FILE="_incbodyline.asp" --------------------------------------------------------- -->
<!--#INCLUDE FILE="_incnav.asp" -------------------------------------------------------------- -->

    <br>
    <table bordercolor="<%= BaseColor %>" cellpadding="0" cellspacing="0" border="0">
      <tr><td><table><tr><td>

<% ' Put Code for top page activity here.
'    response.write "hello"
'    response.end

     Up = strUpNav

'    Validation and System Messages
'    ------------------------------------------------------------------------

      bLstErr = (Session("valCount") > 0 AND Session("valMessage") > "")
  if (bLstErr) then
%>
        <table width="1" cellpadding="0" cellspacing="0">
          <tr><td>&nbsp;&nbsp;&nbsp;</td><td><hr align="left" color="red" width="400" size="1"></td></tr>
          <tr><td>&nbsp;&nbsp;&nbsp;</td><td><font color="red" size="3"><b>Violated Validation Rules: </b></font></td></tr>
          <tr><td>&nbsp;&nbsp;&nbsp;</td><td><font color="red" size="3"><b><%= Session("valMessage") %></b></font></td></tr>
          <!-- tr><td>&nbsp;&nbsp;&nbsp;</td><td><hr align="left" color="red" width="400" size="1"></td></tr -->
        </table>
<%  Else
      bMsg    = (5 < len(Session("Where4List"))) and (Process = "LIST" or (Process = "FORM" and Session("ID") = -1))
  if (bMsg) then
%>
        <table>
          <!-- tr><td>&nbsp;&nbsp;&nbsp;</td><td><hr align="left" color="red" width="400" size="1"></td></tr -->
          <tr><td><font color="maroon" size="2">Filter: <%= Session("Where4List") %></font></td></tr>
          <!-- tr><td>&nbsp;&nbsp;&nbsp;</td><td><hr align="left" color="red" width="400" size="1"></td></tr -->
        </table>
<%    End If
      bMsg    = (Session("Message") > "")
  if (bMsg) then
%>
        <table >
          <!-- tr><td>&nbsp;&nbsp;&nbsp;</td><td><hr align="left" color="red" width="400" size="1"></td></tr -->
          <tr><td><font color="maroon" size="2"><b><%= Session("Message") %></b></font></td>          </tr>          <tr>
          <!-- tr><td>&nbsp;&nbsp;&nbsp;</td><td><hr align="left" color="red" width="400" size="1"></td></tr -->
        </table>
<%    End If
      End If

      Session("Message") = ""


' ==========================================================================================
' END HTML HEADER FOR EVERY PAGE
' ==========================================================================================

' if (strTableName <> "tUNKNOWN") then                                   ' .(40728.01.1
  if (strTableName =  "tUNKNOWN") then                                   '*.(40621.01.5  .(40728.01.1
      response.end                                                       '*.(40621.01.6  .(40728.01.1
      end if                                                             '*.(40621.01.7  .(40728.01.1.
  if (bDebug) then response.write "[1.0] Process is " & ucase(Process) & "<br>"

' Do Processes
' ------------------------------------------------------------------------

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
      case "ADD NEW RECORD"
                     response_redirect 2.21, RootFileName & "?ID=0", ""
      case "FIRST":  response_redirect 2.22, RootFileName & "?Page=1", ""
      case "PREV" :  response_redirect 2.23, RootFileName & "?Page=" & session("nPage") - 1, ""
      case "NEXT" :  response_redirect 2.24, RootFileName & "?Page=" & session("nPage") + 1, ""
      case "LAST" :  response_redirect 2.25, RootFileName & "?Page=" & session("nPageCount"), ""
      case "SHOW" :  response_redirect 2.26, RootFileName,  ""
      case "FILTER": response_redirect 2.27, RootFileName & "?ID=-1", ""
      case "REPORT":
				     Session("ReportName") = aNiceTableName & " Report"
                     session("SQLStrReport") = Session("SQLStr")
                     response_redirect 2.28,"_generatereport.asp", ""
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
    response.write "[2.4] CleanSQL: " & CleanSQL(aSQL) & "<BR>"
    response.write "[2.4] ConnectionString: "  &  SQLConnectionStr & "<br>"
'   response.end
    End If

       on error resume next: if (cnum(bDeBug) = true) then on error goto 0

  Set pRS = Server.CreateObject("ADODB.Recordset")
      pRS.CursorLocation = 3 ' adUseClient

' Set pRS = Conn.Execute (CleanSQL(aSQL))
      pRS.Open CleanSQL(aSQL), SQLConnectionStr, 3, 3

'      Set pRS  = Conn.Execute(aSQL)

      if (Err.Number <> 0) then
          response.write "<font color=red><b>INVALID SQL STATEMENT: " & aSQL & "</b></font><br>"
          setValidation "","[3.2] SQL: "   & aSQL, ""
          setValidation "","ERROR: " & Err.Description, ""
          response_redirect  3.2, RootFileName & "?ID=" & session("ID"), ""
'         response.end
          end if

If (pRS.EOF) Then

    pRS.Close
Set pRS = Nothing

'   setValidation "","[2.4] SQL: "   & aSQL, ""
'   setValidation "","ERROR: No records found.<br>Please enter your first one now.", ""

If (session("RO") <> "YES") then
    Session("message") = "No Records Found.<br>Please enter your first one now."
    response_redirect 2.21, RootFileName & "?ID=0", ""
  else
    Session("message") = "No Records Found."
    response.redirect "_message.asp"
  end if

'   response_redirect 2.21, RootFileName & "?ID=0", ""
'   response_redirect 2.4, "incmessage.asp", ""
'   response.end

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

' FORMAT LIST BUTTON BAR
' ------------------------------------------------------------------------
  function displayListBar

        If (usesMenus) then %>
            <tr>
              <td bgColor="<%=BaseColor%>" align="left" valign="middle" colspan="20">
              <%
           If (ucase(left(Session("RO"),1)) <> "Y") then  %>
              <input type="submit" value="Add New Record" name="btn" <%'= aBtnOff %>><%
              End If %> <input type="submit" value="First" name="btn">
              <input type="submit" value="Prev" name="btn">
              <input type="submit" value="Next" name="btn">
              <input type="submit" value="Last" name="btn">&nbsp;
              <input type="submit" value="Report" name="btn">&nbsp;
              <input type="submit" value="Filter" name="btn">&nbsp;
            <%If false then %>
              <!--input type="submit" value="Show" name="btn"> &nbsp;<input type="text" value="<%= session("ShowRecords") %>" size="3" name="ShowRecords"-->
              <font size="2" color="#FFFFFF">&nbsp;<%= "   " & nPage & " of " & nPageCount & " pages" %>
              </font>
              <%End if%>
              </td>
            </tr>   <%
            End If

            end function
%>
        <!-- TOP OF LIST -------------------------------------------------------- -->

        <table border="1" bordercolor="<%=BorderColor%>" cellspacing="0" style="border-collapse: collapse" cellpadding="0">
          <form method="POST" action="<%= RootFileName & iif(bDebug, "?debug","") %>">

            <!-- TOP LIST BAR -------------------------------------------------------- -->

            <%   displayListBar  %>

            <!-- COLUMN HEADINGS ------------------------------------------------------------ -->
            <%
                response.write aFill & "<TR VALIGN='top' ALIGN='left'>"

'               response.write aFill & "   <td align='left' valign='top' ><font size='2'><b><a href=" & RootFileName & "?Sort=" & pRS(strIDField) & ">" & pRS(strIDField) & "</a></b></font>&nbsp;</td>"

                For Each pFld in pRS.Fields
'                   If (0 = instr(1, strFields4Sys & ",", pFld.Name & ",", 1)) then
'                   If (0 = instr(1, strFields4Sys,       pFld.Name, 1)) then
                    If ( not inFlds( strFieldsNot4List,   pFld.Name)) then
                        If ucase(right(pfld.Name,2)) = "ID" then
                            response.write aFill & "   <td align='left' valign='top' ><font size='2'><b> &nbsp; </b></font>&nbsp;</td>"
                        Else
                            response.write aFill & "   <td align='left' valign='top' ><font size='2'><b><a href=" & RootFileName & "?Sort=" & pFld.Name & ">" & pFld.Name & "</a></b></font>&nbsp;</td>"
                        End If
                        end If
                    Next
                response.write aFill & "</TR>"   %>

            <!-- BEG OF LIST FIELDS --------------------------------------------------------- -->
            <%

' Position recordset to the page we want to see
' ------------------------------------------------------------------------

         on error resume next: if (cnum(bDeBug) = true) then on error goto 0                  ' .(40621.02.1

            pRS.AbsolutePage = nPage

            bBookMarks = (Err.Number <> 3251)                                                 ' .(40621.02.2
        if (bBookMarks = False) then                                                                ' .(40621.02.3
'           response.write "<font color=red><b>TABLE DOES NOT SUPPORT BOOKMARKS</b></font>"   ' .(40621.02.4
'           response.write "<br>"                                                             ' .(40621.02.5
            end if                                                                            ' .(40621.02.6

' Loop through records until it's a next page or End of Records
' ------------------------------------------------------------------------

'    Do While Not (pRS.EOF OR   pRS.AbsolutePage <> nPage                 )                   '*.(40621.02.7
     Do While Not (pRS.EOF OR ((pRS.AbsolutePage <> nPage) AND bBookMarks))                   ' .(40621.02.8

        aIDfld  = pRS.Fields(0).Name
        aIDnum  = pRS.Fields(0).Value:
        'If (isnull(pRS("MarkedCompleteBY")) or ("" + trim(pRS("MarkedCompleteBY")) = "")) AND session("ROLE") <> "RO" Then
         if not usesmarkcomplete AND session("ROLE") <> "RO" Then 
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

            <!-- END OF LIST FIELDS --------------------------------------------------------- -->

            <tr vAlign="top">
              <td bgColor="#F5F5F5" height="4" colspan="3"></td>
            </tr>

            <!-- BOTTOM LIST BAR ----------------------------------------------------- -->

            <%   displayListBar  %>

            <!-- END OF BOTTOM LIST BAR ----------------------------------------------------- -->
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

' FORMAT LIST BUTTON BAR
' ------------------------------------------------------------------------
  function displayFormBar()

        If (usesMenus) then %>
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
          <%  if (usesMarkedComplete) then %>
              <input type="submit" value="Validate"         name="btn" title="Validate"         <%'= aBtnOff %>>
          <%  end if %>
          <%' if (instr("ADM,POC", session("Role")) > 0)    then %>
              <input type="submit" value="Delete"           name="btn" title="Delete"           <%'= aBtnOff %>>
              <input type="submit" value="List"             name="btn" title="List">&nbsp;&nbsp;
              <%
            ' end if  '-- ADM or POC
          end if   '-- RO                                        %>
        <!--  <input type="submit" value="Report"           name="btn" title="Report">  -->
             <%
              if (request("RO") = "ON" AND usesMarkedComplete and session("ROLE") <> "RO")   then %>
                  <input type="submit" value="Unmark Complete"    name="btn" title="Unmark Complete"    <%'= aBtnOff %>>
              <% end if
              if (request("RO") = "ON")   then %>
                  <input type="submit" value="Return"             name="btn" title="Return"             <%'= aBtnOff %>>
              <% end if %>
          <% end if %>
              </td>
            </tr>   <%
        end if

            end function

' -----------------------------------------------------------------------------------------
    select case aAction

     case "EDIT", "NEW", "FILTER"         ' [3.1]

' --------------------------------------------------------------------------------
'             response.write "usesMarkComplete: " & usesMarkedComplete & "<br>"
'             response.write "request(RO): " & request("RO")& "<br>"
'             response.end
%>
        <!-- TOP OF FORM -------------------------------------------------------- -->

        <table border="1" bordercolor="<%= BorderColor %>" cellspacing="0" style="border-collapse: collapse" cellpadding="0">

          <form method="POST" action="<%= RootFileName %>?ID=<%= session("ID") & iif(bDebug, "&debug","") %>">

            <!-- TOP FORM BAR -------------------------------------------------------- -->

            <%   displayFormBar  %>

            <tr vAlign="top">
              <td bgColor="#F5F5F5" height="4" width="1%" colspan="3"></td>
            </tr>

            <!-- BEG OF FORM FIELDS --------------------------------------------------------- --><%

' Display Fields for One Record
' ------------------------------------------------------------------------

            Select Case aAction '
              Case "NEW":    aSQL = "SELECT " & fmtSQL_Flds( strFields4Add   ) & " FROM [" & strTableName & "]"  ' .(40426.01.3
              Case "EDIT":   aSQL = "SELECT " & fmtSQL_Flds( strFields4Edit  ) & " FROM [" & strTableName & "]"  ' .(40426.01.4
              Case "FILTER": aSQL = "SELECT " & fmtSQL_Flds( strFields4Filter) & " FROM [" & strTableName & "]"  ' .(40426.01.5
                end Select

            aSQL = aSQL & " WHERE " & strIDField & " = " & Session("ID")

        if (bDebug = true) then response.write "[3.2] SQL: " & aSQL & "<br>"

        Set pRS  = getRS( aSQL, pFields)

%><!--#INCLUDE FILE="_formX.vbs" --><%

    aPath = Request.ServerVariables("PATH_TRANSLATED"): aPath = left(aPath, instrrev(aPath, "\"))

    aForm = displayCustomForm(  aPath & strCustomForm, pFields, pRS, (session("RO") <> "YES"))

'   aForm = replace( aForm, "{Status}", lookup("Active,     In Review 1,In Review 2,In Review 3,In Review 4,In Review 5,Complete,   ", pRS("Status"), 11))

if (aForm > "") then ' Do this because we have a custom form

    response.write aForm

  else               ' Do this because we DONT'T have a custom form

'     For Each pFld in pRS.Fields
      For Each pFld in pFields

          If ( not inFlds( strIDField & ", " & iif(aAction = "FILTER", strFieldsNot4List, strFields4Sys), pFld.Name)) then
'         --------------------------------------------------------------------------------------------------------------------

                response.write aFill & "<TR>"

            if (bNewRec = true ) then aVal = ""
            if (bNewRec = false) then aVal = pRS( pFld.Name )
            if (bLstErr = true ) then aVal = resetLastValue( pFld.Name, aVal)

                bAddFld = not (pFld.Attributes and &H00000040)                             ' NOT NULL, ie. Required
                bAddFld = bAddFld OR 0 < instr( 1, strFields4Add, pFld.name, 1)            ' OR in Fields4Add List

'           if (bAddFld AND aAction = "NEW") then                                          ' should it be this
'               if (aAction = "NEW" and 0 = instr( 1, strFields4Add, pFld.name, 1)) then   ' it was this

'           Fields for New Record being Added / Inserted
'           -----------------------------------------------
            if (bAddFld AND aAction = "NEW") then                                          ' but what about custom fields

                        aStr =        SetCustomFields(         pFld, pRS ) & ""
                if (    aStr = ""  )  then
                        aStr = "  " & SetRequiredHiddenFields( pFld, pRS ) & ""
                    if (aStr = "  ")  then
                        aStr = "  " & displayField(            pFld, aVal, aHTML) & ""
                        end if
                    end if

'           Fields for Old Record being Edited / Updated
'           -----------------------------------------------
              else
                        aStr =        SetCustomFields(         pFld, pRS) & ""
                    if (aStr = ""   ) then
                        aStr = "  " & displayField(            pFld, aVal, aHTML)
                        end if
                end if
          ' ------------------------------------
          ' Fields for New or Old Records

                if (aFill <> left( aStr, len( aFill) )) then aStr = aFill & aStr
                response.write            aStr
                response.write         msgIfError( pFld.Name )                           ' Err Msg for Field
                response.write aFill & "</TR>"


        End If  ' OK to display Field
'       -------------------------------------------------------------------------

    Next

  end if ' Do this because we don't have a custom form

     on error resume next

        pRS.Close
    Set pRS = Nothing

     on error goto 0

 %>
            <!-- END OF FORM FIELDS --------------------------------------------------------- -->

            <tr vAlign="top">
              <td bgColor="#F5F5F5" height="4" colspan="3"></td>
            </tr>

            <!-- BOTTOM FORM BAR ----------------------------------------------------- -->

            <%   displayFormBar  %>

          </form>
        </table>
        <%
        clrValidation 0    ' After Form is displayed for NEW and EDIT

' -----------------------------------------------------------------------------------------

'    case "SAVE CHANGES", "SAVE & CLOSE"                      ' [4.0]  '*.(40707.03.1
'    case "SAVE CHANGES", "SAVE & CLOSE", "VALIDATE"          ' [4.0]  '*.(40707.03.1
     case "SAVE CHANGES", "SAVE & CLOSE", "VALIDATE", "LIST"  ' [4.0]  ' .(40707.03.1

    if (left(session("RO"),1) <> "Y") then                             ' .(40707.03.5

    if (session("ID") = 0) then  ' Insert New Record

' ---------------------------------------------------------------------------------

' Validate Input for Inserts only
' ------------------------------------------------------------------------

        Check4Duplicates Request           '[4.1]

'       aSQL = "SELECT TOP 1 " &              strFields4Add  & ", CreatedBy, CreatedAt FROM "  & strTableName        '*.(40426.01.7
        aSQL = "SELECT TOP 1 " & fmtSQL_Flds( strFields4Add) & ", CreatedBy, CreatedAt FROM [" & strTableName & "]"  ' .(40426.01.7

    if (bDebug = true) then response.write "[4.1] SQL: " & aSQL & "<br>"

     on error resume next: if (cnum(bDeBug) = true) then on error goto 0

    Set pRS  = Conn.Execute(aSQL)

    if (Err.Number <> 0) then
        response.write "<font color=red><b>INVALID SQL STATEMENT: " & aSQL & "</b></font><br>"
        setValidation "","[3.2] SQL: "   & aSQL, ""
        setValidation "","ERROR: " & Err.Description, ""
        response_redirect  3.2, RootFileName & "?ID=" & session("ID"), ""
'       response.end
        end if

'       chkValidation pRS                 ' [4.2] Checks for NOT NULL, Dates and Numbers
'       Check4Validation Request          ' [4.5] Table Specific Checks

    if (bDebug = true) then response.write "[4.2] Validation Errors: " & session("valCount") & "<br>"

    if (session("valCount") > 0) then
        response_redirect  4.2, RootFileName & "?ID=" & session("ID"), ""
        end if

' Insert New Record
' ------------------------------------------------------------------------
        dNow = now()

    if (strFields4Add = "*") then
'       strFields4Add = getAllFieldsBut(pRS, strIDField & "," & strFields4Sys & ",CreatedBy,CreatedAt")
        strFields4Add = getAllFieldsBut(pRS, strIDField & "," & strFields4Sys & ",CreatedBy,CreatedAt,ChangedBy,ChangedAt")                 ' .(40707.06.1
        end if

'                  aSQL = " INSERT "  & strTableName &  " (" &              strFields4Add  & ", CreatedBy, CreatedAt)"                      '*.(40426.01.8
'                  aSQL = " INSERT [" & strTableName & "] (" & fmtSQL_Flds( strFields4Add) & ", CreatedBy, CreatedAt)"                      '*.(40707.06.1 .(40426.01.8
                   aSQL = " INSERT [" & strTableName & "] (" & fmtSQL_Flds( strFields4Add) & ", CreatedBy, CreatedAt,ChangedBy,ChangedAt)"  ' .(40707.06.1 .(40426.01.8
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
            aSQL = aSQL &        c   & fmtFValue( pRS("ChangedBy"), Session("ChangedBy"))
            aSQL = aSQL &       ", " & fmtFValue( pRS("ChangedAt"), dNow                )
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
        Session("Message") = "Record Added Successfully."  'Record ID " & Session("ID") & " Added."

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

'       chkValidation pRS                 ' [4.5] Checks for NOT NULL, Dates and Numbers   '*.(40707.05.1
'       Check4Validation Request          ' [4.5] Table Specific Checks                    '*.(40707.05.2

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

        Session("Message") = "Record Saved"   'Record ID " & Session("ID") & " Saved."

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

    end if                                                                                ' .(40707.03.6

' Return to FORM or LIST
' ------------------------------------------------------------------------

'   if ((aAction = "SAVE & CLOSE") and session("valCount") = 0) then                      ' .(40707.04.2
'   if ((aAction = "SAVE & CLOSE" or aAction = "LIST") and session("valCount") = 0) then  ' .(40707.04.2

    select case aAction                                                                   ' .(40707.04.3
      case "SAVE & CLOSE", "LIST"                                                         ' .(40707.04.4

        response_redirect  4.91, RootFileName, ""

'     else                                                                                '*.(40707.04.5
      case "SAVE CHANGES"                                                                 ' .(40707.04.5

        response_redirect  4.92, RootFileName & "?ID=" & Session("ID"), "FORM"
'       end if                                                                            '*.(40707.04.5

' -----------------------------------------------------------------------------------------

      case "VALIDATE"                      ' [5.0]

    set pRS  =  Server.CreateObject( "ADODB.Recordset" )
'       pRS.Open "SELECT " &              strFields4Edit  & " FROM "  & strTablename &  " WHERE " & strIDField & " = " & Session("ID"), Conn  '*.(40426.02.2
        pRS.Open "SELECT " & fmtSQL_Flds( strFields4Edit) & " FROM [" & strTablename & "] WHERE " & strIDField & " = " & Session("ID"), Conn  ' .(40426.02.2

        chkValidation pRS

        Check4Validation pRS

    if (session("valCount") = 0) then
        response.redirect "_formvalidate.asp"
        Session("Message")  = "All Fields in Record ID " & Session("ID") & " Are Valid."
        end if

        response_redirect 5, RootFileName & "?ID=" & Session("ID"), ""

        end select                                                                        ' .(40707.04.6
' -----------------------------------------------------------------------------------------

     case "DELETE"                        ' [6.0]

'       aSQL = "DELETE "  & strTablename &  " WHERE " & strIDField & " = " & Session("ID")    '*.(40426.02.3
        aSQL = "DELETE [" & strTablename & "] WHERE " & strIDField & " = " & Session("ID")    ' .(40426.02.3

    if (bDebug = true) then
        response.write "[6.0] SQL: " & aSQL & "<br>"
      else
        Conn.Execute(aSQL)
        end if

        Session("Message") = "Record Deleted" 'ID " & Session("ID") & " Deleted. " & iif(bDebug, "<font color=red>(Not if debugging)</font>", "")

        response_redirect 6.0, RootFileName, ""

' -----------------------------------------------------------------------------------------

     case "REPORT"                        ' [7.0]

        Session("ReportName") = aNiceTableName & " Report"
        Session("SQLStrReport") = Session("SQLStr")
        response_redirect 7.0, "_generatereport.asp", ""

' -----------------------------------------------------------------------------------------

     case "UNMARK COMPLETE"                 ' [3.8]

        Session("Process") = "LIST"   ' "UNMARK COMPLETE"
        SQLStrtmp = "UPDATE tdatacall Set MarkedCompleteat = '', MarkedCompleteBy = '' WHERE datacallID = "& Session("ID")
'       response.write SQLStrtmp
'       response.end
        Conn.Execute(SQLStrtmp)
        response_redirect 3.75, RootFileName, ""

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
    set pRS1  =   Server.CreateObject( "ADODB.Recordset" )
'          bldFilter pRS
          Where4List = Session("Where4List")  ' Save old filter
'         pRS.Open "SELECT TOP 1 " &              strFields4Filter  & " FROM "  & strTablename &  " WHERE " & strIDField & " = " & Session("ID"), Conn  '*.(40426.02.4
          pRS.Open "SELECT TOP 1 " & fmtSQL_Flds( strFields4Filter) & " FROM [" & strTablename & "] WHERE " & strIDField & " = " & Session("ID"), Conn  ' .(40426.02.4
          Session("Where4List") = bldFilter( pRS )
        sqlstr =  "SELECT count(*) as reccount FROM [" & strTablename & "] WHERE " & Session("Where4List")
        response.write sqlstr
          pRS1.Open "SELECT count(*) as reccount FROM [" & strTablename & "] WHERE " & Session("Where4List"), Conn  ' .(40426.02.4
          If pRS1("reccount") = 0 then
             Session("Message") = "No records found for Filter: " & Session("Where4List")
             Session("Where4List") = Where4List ' Use old filter
          end if
        pRS1.Close
    set pRS1 = Nothing
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

     case "RETURN"                        ' [7.8]

        response_redirect 7.8, RootFileName, ""

' -----------------------------------------------------------------------------------------

     case else

        Session("Message") = "Form Action System Error: Process = " & aAction
        response_redirect 7.9, RootFileName & "?ID=" & session("ID"), ""

      end select ' FORM aAction


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
          <form method="POST" action="<%= RootFileName %>?ID=<%= session("ID") & iif(bDebug, "&debug","") %>&Process=MARK COMPLETE">
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
                    <% If (usesMarkedComplete AND instr("POC",Session("Grouplist")) > 0 OR instr("ADM",Session("Grouplist")) > 0 OR Session("PersonLastName") = "Schinner") Then %>
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

    If (usesMarkedComplete) then
'       aSQL = "UPDATE "  & strTableName &  " Set MarkedCompleteBy = '"& Session("ChangedBy") &"', MarkedCompleteAt = '"& Now() & "' WHERE " & strIDField & " = " & Session("ID")  '*.(40426.02.5
        aSQL = "UPDATE [" & strTableName & "] Set MarkedCompleteBy = '"& Session("ChangedBy") &"', MarkedCompleteAt = '"& Now() & "' WHERE " & strIDField & " = " & Session("ID")  ' .(40426.02.5
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
'  CASE ELSE
' ==========================================================================================

        Session("Message") = "System Process Error"
        response_redirect 10, RootFileName, ""

   END SELECT


' ------------------------------------------------------------------------
' End Processes

'     end if                                                             ' .(40728.01.2.

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
       dim pFld, aFld

    for i = 0 to pRS.fields.count - 1

'       If (not inFlds( strIDField & ", " & strFields4Sys,                            pRS.Fields(i).Name)) then
'       If (not inFlds( strIDField & ", " & strFields4Sys & ", ChangedBy, ChangedAt", pRS.Fields(i).Name)) then   '*.(40619.03.1
        If (not inFlds( strIDField & ", " & strFields4Sys & ", CreatedBy, CreatedAt", pRS.Fields(i).Name)) then   ' .(40619.03.1

        set pFld = pRS.Fields(i)
            aFld = SavCustomFields(pFld, request(pFld.Name)) & ""
        if (aFld = "") then

            aFld = fmtSQL_FldEqRequestVal( pFld )
            end if
'           response.write "FLD: '" & aFld & "'<br>"

            aSQL = aSQL & ", " & aFld

            end if
        next

        if (inFlds( strFields4Sys, "ChangedBy")) then                               ' .(40619.03.2
            aSQL = aSQL & ", " & "ChangedAt = '"        &  now() & "'"
'           aSQL = aSQL & ", " & "ChangedBy = '"        &  replace(Session("ChangedBy"), "'", "''") & "'"
            aSQL = aSQL & ", " & "ChangedBy = " & fmtSQL_Val( pRS("ChangedBy"), Session("ChangedBy"))
            end if                                                                  ' .(40619.03.3

        If (usesMarkedComplete) then
            aSQL = aSQL & ", " & "MarkedCompleteAt = "  & "NULL"
            aSQL = aSQL & ", " & "MarkedCompleteBy = "  & "NULL"
            End If

'           aSQL = "UPDATE "  & strTableName &  " SET " & mid(aSQL, 3) & " WHERE " & strIDField & " = " & session("ID")  '*.(40426.03.1
            aSQL = "UPDATE [" & strTableName & "] SET " & mid(aSQL, 3) & " WHERE " & strIDField & " = " & session("ID")  ' .(40426.03.1

'       bDebug = true
    if (bDebug = true) then
        response.write "[4.7] SQL: " & aSQL & "<br>"
        response.end
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

<!--#INCLUDE FILE="_incfooter.asp"-------------------------------------------------------------->

</body>

</html>