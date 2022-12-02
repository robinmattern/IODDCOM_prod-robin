<%
          notInASP = not(isObject(response))
      if (notInASP) then  ' Run locally as vbs script

     dim  conn, aRS

     set  conn  = setCS("gsa_htw", "htw", "")
     set  pRS   = getMTRecordset("tapp")

'         aPath = "D:\Users\Robin\Work\ICE\AAP\"
'         aPath = "D:\Inetpub\wwwroot\secureaddress\govwebs\fido\dhs\aap\"
'         aPath = "D:\Inetpub\wwwroot\secureaddress\govwebs\fido\gsa\htw\robin\"
          aPath = "D:\Inetpub\wwwroot\secureaddress\govwebs\fido\gsa\htw\"

          strCustomForm     =  aPath & "form_app.htm"
          
          bInp              =  true
          bDoItHere         =  true
          bShowCustomFields =  true
          
          pRS("Status")          = "Rejected by Legal"
          pRS("AdministratorBy") = "Robin Mattern"

          aForm = displayCustomForm( strCustomForm, pRS.Fields, pRS, bInp)

       else

'         aForm = displayCustomForm( strCustomForm, pFields,    pRS, bInp)   ' done by calling ASP file

          end if

' ------------------------------------------------------------------------

function  getFields4CustomForm_( pFlds, pRS, bInp) 

     dim  mFlds2()
'  redim  mFlds2( ubound( mFlds1 )    )
   redim  mFlds2(         pFlds.count )

          mFlds2(  2) = setFld( "text"    , "Region"                     ,    12, pRS,   bInp)
          mFlds2(  1) = setFld( "text"    , "ControlNumber"              ,    12, pRS,   bInp)
          mFlds2( 13) = fmtSel(             "Status"                     ,        pRS,   bInp)
          mFlds2( 14) = setFld( "textarea", "AdministratorComment"       ,    50, pRS,   bInp)
          mFlds2(  3) = setFld( "text"    , "AdministratorBy"            ,    56, pRS,   bInp)
          mFlds2(  4) = setFld( "text"    , "AdministratorAt"            ,    12, pRS,   bInp)
          mFlds2( 15) = setFld( "textarea", "GeneralCounselComment"      ,    50, pRS,   bInp)
          mFlds2(  5) = setFld( "text"    , "GeneralCounselBy"           ,    56, pRS,   bInp)
          mFlds2(  6) = setFld( "text"    , "GeneralCounselAt"           ,    12, pRS,   bInp)
          mFlds2( 16) = setFld( "textarea", "LegalComment"               ,    50, pRS,   bInp)
          mFlds2(  7) = setFld( "text"    , "LegalBy"                    ,    56, pRS,   bInp)
          mFlds2(  8) = setFld( "text"    , "LegalAt"                    ,    12, pRS,   bInp)
          mFlds2( 17) = setFld( "textarea", "GSAFleetComment"            ,    50, pRS,   bInp)
          mFlds2(  9) = setFld( "text"    , "GSAFleetBy"                 ,    56, pRS,   bInp)
          mFlds2( 10) = setFld( "text"    , "GSAFleetAt"                 ,    12, pRS,   bInp)
          mFlds2( 18) = setFld( "textarea", "RegionComment"              ,    50, pRS,   bInp)
          mFlds2( 11) = setFld( "text"    , "RegionBy"                   ,    56, pRS,   bInp)
          mFlds2( 12) = setFld( "text"    , "RegionAt"                   ,    12, pRS,   bInp)

          getFields4CustomForm_ = mFlds2
     end  function

' ------------------------------------------------------------------------

function  displayCustomForm( aFile1, pFlds, pRS, bInp ) 
     dim  aForm1, mFlds1, mFlds2, aStr, i          
     
      if (not fileExists(aFile1)) then 
          displayCustomForm = ""
          exit function
          end if

          aForm1 = getFile( aFile1 )
          mFlds1 = getInpTags( aForm1)

      if (bShowCustomFields) then

          sayMsg "Fields in Original Form: " & aFile1 
          
                   shoFlds mFlds1, 0, 0
'                  shoFlds mFlds2, 5, 0

          sayMsg vbCrLf & vbCrLf & "Field Values in Database Table: " & vbCrLf: i = 0

      for each pFld in pFlds
          i = i + 1
          sayMsg right("   " & i, 3) & ": " & pFld.Name & " = " & pRS(pFld.Name)  
          next

          end if

'     For Each pFld in pFlds
'
'         aStr =     SetCustomFieldValue( aFld, pRS(aFld)) & ""
'         aStr =     SetCustomFieldValue( pFld, pRS      ) & ""
'     if (aStr > "") then
'         pRS( pFld.Name ) = aStr
'         end if
'         
'         Next

      if (cnum(bDoItHere)) then 
          mFlds2 = getFields4CustomForm_( pFlds, pRS, bInp) 
        else  
          mFlds2 = getFields4CustomForm(  pFlds, pRS, bInp) 
          end if

      if (bShowCustomFields) then
      
          sayMsg  vbCrLf  & vbCrLf & "Fields in Populated Form Array: mFlds2" 
          shoFlds mFlds2, 0, 0
          end if 
          
          aForm3 = incForm( aForm1, mFlds1, mFlds2, bInp)

      if (bShowCustomFields) then
          sayMsg  vbCrLf  & vbCrLf & "Fields in Populated Form String: aForm3" 
          shoFlds getInpTags( aForm3 ), 0, 0
          
          end if

          displayCustomForm = aForm3
      
      end function
         
' --------------------------------------------------------------------------------

   function shoFlds( mFlds, nFld, nFlds)
        dim aStr
            sayMsg ""
'           sayMsg "nFlds: " & ubound(mFlds)
        if (0 = nFld)  then nFld = 1
        if (0 = nFlds) then nFlds = ubound(mFlds) - (nFld - 1)     
            sayMsg "nFlds: " & (nFld + (nFlds - 1))
        if (ubound(mFlds) < nFld + (nFlds - 1)) then nFlds = ubound(mFlds) - (nFld - 1)
        
        for i = nFld to nFld + (nFlds - 1)
            aStr = mFlds(i)
        if (isObject(response)) then 
            aStr = replace(replace(aStr,"<","&lt;"),">","&gt;")
            end if
            sayMsg(right("   " & i, 3) & ": " & aStr)
            next
        end function

' -----------------------------------------------------------------------

'  function incForm( aFile1, mFlds2, bInp)
'       dim aForm1,  mFlds1, aFld1,  aFld2, nFlds, i

   function incForm( aForm1, mFlds1, mFlds2, bInp)
        dim aFld1,  aFld2, i

          nFlds  =  ubound(mFlds1)
      if (nFlds  >  ubound(mFlds2) ) then 
          nFlds  =  ubound(mFlds2)
          end if

        for i = 1 to nFlds
            aFld1  = mFlds1(i)
            aFld2  = mFlds2(i) & ""

        if (aFld2  > "" or bInp = false) then

            aFld2  = fmtFld( aFld1, aFld2)

'           sayMsg( right(" " & i,2) & ": " & aFld1 & " <=9 " & aFld2)

        if (bInp and "<" = left(aFld2,1)) then 
            aFld2  =  setAtr("class", aFld2, "InpFld", "" )
            end if

            aForm1 = replace( aForm1, aFld1, aFld2)

            end if
            next

'           sayMsg(   aForm1 )
'           response.end

            m =  getInpTags( aForm1)
'           shoFlds  m, 45, 1
'           shoFlds  m, 1, 45
'           shoFlds  m, 41, 2

            incForm = aForm1
        end function

' --------------------------------------------------------------------------------

function fmtFld( byVal aFld1, byVal aFld2)
     dim aType, aTag, aVal

     if (0 < instr(aFld2, "<")) then

         aFld1  =  setAtr("name",    aFld1,  aFld2,   ""  )
'        aFld1  =  setAtr("class",   aFld1, "InpFld", ""  )

         aTag   = "textarea"
     if (aTag   =  lcase(mid(aFld1,2,len(aTag)))) then
         aVal   =  getStr(aFld2, ">", "<", nBeg, nLen, false)

'        sayMsg(  "textarea: " & aFld2)
'        sayMsg(   aFld2 & ": " & nBeg & ", " & nLen & " - '" & aVal & "'"):
'        sayMsg(   take( 10 + nBeg - 1, "") & "^" & take( nLen - 1, "") & "^")
'        sayMsg(  "textarea: " & aFld1 )

                   getStr aFld1, ">", "<", nBeg, nLen, false
         aFld2  =  left(  aFld1, nBeg) & aVal & mid(aFld1, nBeg + nLen - 1)

'        sayMsg(  "textarea: " & aFld2)
         end if

         aTag   = "select"
     if (aTag   =  lcase(mid(aFld1,2,len(aTag)))) then
'        aFld2  =  setAtr("type",    aFld1, aFld2, aType  )
'        sayMsg("aFld2: " & aFld2) 
'        aFld2  =  replace(aFld1, "'" & aVal & "'", "'" & aVal & "' selected") 
'        aFld2  =  aFld1
'        sayMsg("aFld2: " & aFld2) 
         end if
         
         aTag   = "input"
     if (aTag   =  lcase(mid(aFld1,2,len(aTag)))) then
         aType  = "text"
         aFld1  =  setAtr("type",    aFld1, aFld2, aType  )
     
'        sayMsg("aFld1: " & aFld1) 
     if (aType  = "text") then
         aFld1  =  setAtr("size",    aFld1, aFld2, "10"   )
         end if
         
     if (aType  = "radio" or aType = "checkbox") then
         aFld1  =  replace(aFld1,   " checked ", " checked=""true""" )
         aFld2  =  replace(aFld2,   " checked ", " checked=""true""" )
         aFld1  =  setAtr("checked", aFld1, aFld2, "false")
         aFld1  =  replace(aFld1,   " checked=""true""",  " checked ")
         aFld1  =  replace(aFld1,   " checked=""false""", "         ")
'        aFld1  =  setAtr("size",    aFld1, "size=""null""", "")
         aFld1  =  setAtr("size",    aFld1, "null", "")
         end if
         aFld2  =  setAtr("value",   aFld1, aFld2, ""     )
         end if

         end if ' aFld2 is not MT

         fmtFld = aFld2
     end function

' --------------------------------------------------------------------------------

Function setAtr( aAtr, byVal aStr1, byVal aStr2, aDefault)
     Dim aVal1,  aVal2, nPos1

         aVal1 = getStr(aStr1, aAtr & "=""", """", nPos1, 0, False)
         aVal2 = getStr(aStr2, aAtr & "=""", """", 0,     0, False)
     if (aVal2 = "" and 0 = instr(aStr2, "<")) then aVal2 = aStr2

         setAtr = aStr1
     If (aVal1 > "" And aVal2 = "") Then Exit Function
'    if (               aVal2 > "") Then aVal2 = aVal2
     If (aVal1 = "" And aVal2 = "") Then aVal2 = IIf(aDefault > "", aDefault, "")
     If (               aVal2 = "") Then Exit Function

     If (nPos1 = 0) Then
'        nPos1 = Len(aStr1) + 1: aVal1 = ""
         nPos1 = instr(aStr1,">") + 1: aVal1 = ""
     if (nPos1 > 1) then 
         aStr1 = Left(aStr1, nPos1 - 2) & " " & aAtr & "="""">" & mid(aStr1,nPos1)
         end if
'        if ("<TEXTAREA" = left(aStr1,9)) then sayMsg "aStr1: " & aStr1 
         End If

         nPos3 = nPos1
         nPos1 = nPos1 + Len(aAtr) + 1
     if (aVal2 = "null") then
'        sayMsg("*: " & aStr1 & " <=8 " & aStr2)
'        sayMsg( take( 3 + nPos1-1, "") & "^")
'        aMsg = "Changing nPos from (" & nPos1 & "," & nPos2 & ") to "
         nPos2 = nPos1 + 1: nPos1 = nPos3 - 1: aVal2 = ""
'        sayMsg( take( 3 + nPos1-1, "") & "^" & take( nPos2 - nPos1, "") & "^")
'        sayMsg( aMsg & "(" & nPos1 & "," & nPos2 & ")")
       else
         nPos2 = nPos1
         end if
'        sayMsg( ucase(aAtr) & "(" & nPos1 & "): " & Left(aStr1, nPos1) & "{}" & Mid(aStr1, nPos1 + Len(aVal1) + 1))

         aStr1 = Left(aStr1, nPos1) &           aVal2         & Mid(aStr1, nPos2 + Len(aVal1) + 1)
'        aStr1 = Left(aStr1, nPos1) & take( 15, aVal2 & """") & Mid(aStr1, nPos2 + Len(aVal1) + 2)

         aDefault = aVal2
         setAtr   = aStr1
     End Function

' --------------------------------------------------------------------------------

function ifCkd( aVal, aStd)
         ifCkd = iif(trim(ucase(aVal & "")) = trim(ucase(aStd & "")), " checked", "        ")
     end function

' --------------------------------------------------------------------------------

function setFld( byVal aTyp, aFld, nWdt, byVal aVal, bInp)
     Dim aStr,   pRS
'        aVal = "Robin"
     if (isObject(aVal)) then
'        sayMsg "aFld: " & aFld & " = " & aVal(aFld)
     Set pRS  =  aVal
         aVal =  pRS(aFld)
         end if

         aTyp = iif(bInp = true,"W","R") & trim(left(lcase(aTyp),5))
'        sayMsg("Type: " & aTyp)

       select case aTyp
         case "Wtext":  aStr = "<INPUT type=""text""     name=""" & take( 35, aFld & """") & " size=""" & nWdt      & """ value=""" & aVal & """>"
         case "Rtext":  aStr = aVal
         case "Whidde": aStr = "<INPUT type=""hidden""   name=""" & take( 35, aFld & """") & " size=""" & nWdt      & """ value=""" & aVal & """>"
         case "Rhidde": aStr = aVal
         case "Wtexta": aStr =   "<TEXTAREA              name=""" & take( 35, aFld & """") & " cols=""" & nWdt      & """>"         & aVal & "</TEXTAREA>"
         case "Rtexta": aStr = aVal
         case "Wradio": aStr = "<INPUT type=""radio""    name=""" & take( 35, aFld & """") &       ifCkd( nWdt, aVal) & " value=""" & aVal & """>"
         case "Rradio": aStr = "<INPUT type=""radio""    name=""" & take( 35, aFld & """") &       ifCkd( nWdt, aVal) & " value=""" & aVal & """>"
         case "Wcheck": aStr = "<INPUT type=""checkbox"" name=""" & take( 35, aFld & """") &       ifCkd( nWdt, aVal) & " value=""" & aVal & """>"
         case "Rcheck": aStr = "<INPUT type=""checkbox"" name=""" & take( 35, aFld & """") &       ifCkd( nWdt, aVal) & " value=""" & aVal & """>"
'        case "Wselec": aStr =   "<SELECT                name=""" & take( 35, aFld & """") &       ifCkd( nWdt, aVal) & "   value=""" & aVal & """></SELECT>"
         case "Wselec": aStr = fmtSel( aFld, pRS, nWdt) 
         case "Rselec": aStr = aVal
         case "W":      aStr =   "<INPUT                 name=""" & take( 35, aFld & """") & " size=""" & nWdt      & """ value=""" & aVal & """>"
         case "R":      aStr = aVal
         case else: sayMsg "setFld[1] *** aTyp, " & aTyp & ", not found: "
           end select

         setFld = aStr

'        sayMsg(aFld & ": " & aStr)
     end function

' --------------------------------------------------------------------------------

Function getStr(aStr, aBeg, aEnd, nBeg, nLen, bIncludeEnds)
         nBeg = InStr(1, aStr, aBeg, 1): nLen = Len(aStr)
     If (nBeg = 0) Then Exit Function
         nLen = InStr(nBeg + Len(aBeg), aStr, aEnd, 1)
     If (nLen = 0) Then Exit Function
         nLen = nLen - (nBeg - Len(aEnd))
     If (bIncludeEnds = True) Then
         getStr = Mid(aStr, nBeg, nLen)
       Else
         getStr = Mid(aStr, nBeg + Len(aBeg), nLen - Len(aBeg + aEnd))
         End If
     End Function
'*\
'#PROG     .---------------------+----+----------------------------------------+
'#ID     A2.38. getFile(aFile)   |PMIO| Read Text File
'#SRC      .---------------------+----+----------------------------------------+
'*/
      Function  getFile(aFile)
           Dim  pFO
                aFile = replace(aFile, "/", "\")
            On  Error resume next
           Set  pFO = makObject("Scripting.FileSystemObject")
'          Set  pFS = pFO.OpenTextFile( aFile, 1)                                                                                                                                       '*.(15.01.1
           Set  pFS = pFO.OpenTextFile( aFile, 1, True, bUnicode):  if (Err) then sayMsg("getFile[] Error " & Err.Number & ": " & Err.Description & vbCrLf & "Opening File: " & aFile)   ' .(15.01.1
                getFile = pFS.ReadAll()
                pFS.Close
           Set  pFS = Nothing
           End  Function
'*\
'#PROG     .---------------------+----+----------------------------------------+
'#ID   A2.3281. fileExists(aFile)|PMIO| See if File Exists
'#SRC      .---------------------+----+----------------------------------------+
'*/
      Function  fileExists(aFile)
           Dim  pFO
                aFile = replace(aFile, "/", "\")
           Set  pFO = makObject("Scripting.FileSystemObject")
                fileExists = pFO.FileExists(aFile) 
           End  Function

' --------------------------------------------------------------------------------
 function take(n,a)
          take = left(a & "                                                                                                ",n)
      end function
' --------------------------------------------------------------------------------

 function cNum(n)
          cNum = 0: if (isNumeric(n)) then cNum = cDBL(n)
      end function
' --------------------------------------------------------------------------------

 function nMin(a,b)
          nMin = cNum(a): if (nMin > cNum(b)) then nMin = cNum(b)
      end function
' --------------------------------------------------------------------------------

 function nMax(a,b)
          nMax = cNum(a): if (nMax < cNum(b)) then nMax = cNum(b)
      end function
' --------------------------------------------------------------------------------

 function iif(a,b,c)
          iif = b: if (a <> true) then iif = c
      end function
' --------------------------------------------------------------------------------

function getInpTags( aForm)
     Dim mFlds1, mFlds3

         mFlds1 = getTags(aForm, "INPUT",   False)

         mFlds3 = getTags(aForm, "SELECT",   True)
            n   = ubound(mFlds1)
      redim preserve mFlds1(n + ubound(mFlds3) )
        for i = 1 to ubound(mFlds3)
            mFlds1(n+i) = mFlds3(i)
            next

         mFlds3 = getTags(aForm, "TEXTAREA", True)
            n   = ubound(mFlds1)
      redim preserve mFlds1(n + ubound(mFlds3) )
        for i = 1 to ubound(mFlds3)
            mFlds1(n+i) = mFlds3(i)
            next

         getInpTags = mFlds1
     end function

' --------------------------------------------------------------------------------

function getTags( aStr, aTag, bEnd)
     Dim pRE, pRC, pIT, mFlds(), i, m, n: i = 0
     Set pRE = New RegExp
         pRE.Pattern     = "<" & aTag & " "
         pRE.IgnoreCase  = True
         pRE.Global      = True
         aEnd = iif(bEnd, "</" & aTag & ">", ">")

     Set pRC =  pRE.Execute(aStr)
   ReDim mFlds( pRC.Count)
     If (pRC.Count > 0) Then
     For Each pIT in pRC
         i = i + 1: n = pIT.FirstIndex + 1
         m = instr( n, aStr, aEnd, 1): 
     if (m = 0) then 
         sayMsg "*** Tag terminator, '" & aEnd & "', Not Found for Tag: " & mid(aStr,n, 100) 
       else  
         mFlds(i) = mid( aStr, n, len(aEnd) + m - n)
'        sayMsg( right(" " & i,2) & "[" & right("    " & n,5) & "] " & mFlds(i))
         end if
         Next
       Else
'        sayMsg("*** " & pRE.Pattern & " not found")
         End If

     Set pRE = nothing
         getTags = mFlds
     end function

' --------------------------------------------------------------------------------

function setCS(aDB, aUID, aPWD)
     if (aUID = "") then aUID = aDB
     if (aPWD = "") then aPWD = "pass" & aUID & "1"
         
     SET Conn = makObject("ADODB.Connection")
         aStr = "PROVIDER=MSDASQL;DRIVER={SQL Server};Server=(local);Database=" & aDB & ";Uid=" & aUID & ";Pwd=" & aPWD
'        sayMsg(aStr)
         Conn.open aStr 
     Set setCS = Conn
     end function

' --------------------------------------------------------------------------------

 function getMTRecordset(aSQL)

      if (0 = instr(ucase(aSQL), "SELECT ")) then
          aSQL = "SELECT * FROM [" & aSQL & "]"
          end if

          aSQL = left(aSQL, instr(ucase(aSQL) & " WHERE", " WHERE"))
          aSQL = replace(aSQL, "SELECT ", "SELECT TOP 1 ")
'         sayMsg "SQL: " & aSQL

      Set pRS = makObject("ADODB.Recordset")
'         sayMsg conn.ConnectionString

'         aSQL = "SELECT * FROM tDataCall WHERE DataCallID = 1"
'         aSQL = "SELECT * FROM tDataCall WHERE DataCallID = 0"
'         aSQL = aSQL & " WHERE 1 = 0"
          pRS.Open aSQL, Conn, 3, 3

'     Set pRS  = Conn.Execute(aSQL)

          pRS.AddNew

'  sayMsg pRS.Fields(1).Name & ": " & pRS.Fields(0).Value

'     For Each pField in pRS.Fields
'         pField.value = null
'         Next

      Set getMTRecordset = pRS

      End Function

' --------------------------------------------------------------------------------

   function lookup( aList, aVal, w) 
        dim n
'       response.write "aList = '" & aList & "' in '" & trim(aVal) & ",' is " & instr(aList, trim(aVal) & ",") & "<br>"
            n = instr(aList, trim(aVal) & ",") 
'       if (n > 0) then n = (n + (w - 1)) / w
        if (n > 0) then n = (n -  1     ) / w
            lookup = n
        end function 
        
' --------------------------------------------------------------------------------

   function shoEm( mFlds)

            sayMsg("")
        for i = 1 to 3
            sayMsg(" " & i & ": " & mFlds(i))
            next

            sayMsg("24: " & mFlds(24))
            sayMsg("25: " & mFlds(25))
            sayMsg("43: " & mFlds(43))
            sayMsg("47: " & mFlds(47))
            sayMsg("51: " & mFlds(51))
            sayMsg("57: " & mFlds(57))
'           sayMsg("60: " & mFlds(60))

'       for i = 24 to 25
'           sayMsg(i & ": " & mFlds(24))
'           next
        end function

' --------------------------------------------------------------------------------

function getSession(aVar)
     If (IsObject(Session)) Then
         getSession = Session(aVar)
       else
         select case aVar
           case "RO":            getSession = iif(bInp, "YES", "NO")
           case "Organization":  getSession = "/DHS/"
           case else:    sayMsg("No Session Var: " & aVar)
             end select
         end if
    end function
'*\
'#PROG     .---------------------+----+----------------------------------------+
'#ID     A2.99. sayMsg(aMsg)     | FIO| Say a Mag
'#SRC      .---------------------+----+----------------------------------------+
'*/
       Function sayMsg(aMsg)
         If (IsObject(Response)) Then
             aMsg = replace(aMsg, "'", "&rsquo;")
'            aMsg = replace("<script> alert('" & replace(aMsg, "'", "\'") & "'); </script>", vbCrLf, "\n")
             aMsg = replace( replace(aMsg, "<", "&lt;"), ">", "&gt;")
             aMsg = replace( aMsg & vbCrLf, vbCrLf, "<br>" & vbCrLf)
             Response.write aMsg 
             Exit Function
             End If
'        On Error Resume Next:           MsgBox         (aMsg)
'        If (Err) then
         If (IsObject(WScript    )) Then WScript.Echo   (aMsg)
'        If (IsObject(Application)) Then Debug.Print    (aMsg)
'            End If
     End Function
'*\
'#PROG     .---------------------+----+----------------------------------------+
'#ID     A2.98. makObject(aObj)  | FIO| Create a COM/ActiveX Object
'#SRC      .---------------------+----+----------------------------------------+
'*/
       Function makObject(aObj)
            If (len(aObj) = 4 + instr(aObj, ".wsc")) Then _
                                                  Set makObject =            GetObject(aObj)
            If (IsObject(WScript    )      ) Then Set makObject = WScript.CreateObject(aObj)
            If (IsObject(Server     )      ) Then Set makObject =  Server.CreateObject(aObj)
'           If (IsObject(Application)      ) Then Set makObject =         CreateObject(aObj)
            If (IsObject(makObject) = false) Then Set makObject =         CreateObject(aObj)
        End Function

' ------------------------------------------------------------------------
' These functions are also in .ASP Include file: _incCommonFunctions.asp
' ------------------------------------------------------------------------

 function fmtSel(  aFld, pRS, bBlank1st)
      Dim nWdt, aTable, aKeyFld, aValFld, aWhere, aSQL, pRS1

          nWdt       =  60            ' Width of <OPTION Value="{%= take( nWdt, aKeyVal) %}">{%= aValue %}</OPTION>
          aFill      =  vbCrLf & ""   ' spaces before <SELECT tag

          aTable     = "tLkUp"
          aKeyFld    = "LkUpValue"
          aValFld    = "LkUpValue"
          aSrtFld    = "LkUpID"
          aWhere     = "WHERE (LkUpType = '" & aFld & "')"

      if (aKeyFld   <>  aValFld) then
          aKeyFld    =  aKeyFld & "," & aValFld
          end if

          aSQL       = "SELECT " & aKeyFld & " FROM " & aTable & " " & aWhere & " ORDER BY " & aSrtFld
'  sayMsg("SQL: " & aSQL)
   
      Set pRS1       =  Conn.Execute( aSQL)

      dim aStr:         aStr   = aFill & "<SELECT name=""" & aFld & """>"
      if (bBlank1st) then 
          aStr       =  aStr   & aFill & "  <OPTION value=""""></OPTION>"
          end if
          aStr       =  aStr   &              fmtOptions( aFill, aKeyFld, aValFld, pRS(aFld), pRS1, nWdt )
          aStr       =  aStr   & aFill & "</SELECT>"

          pRS1.close
      Set pRS1       =  Nothing

          fmtSel     =  aStr
      end function

 function fmtOptions( aFill, aKeyFld, aValFld, aCurVal, pRS, nWdt)
      Dim aKeyVal, aValue, aChecked, aOptions: aOptions = ""

       do while (pRS.EOF = false)
          aKeyVal    =  trim( pRS(aKeyFld) )
          aValue     =  trim( pRS(aValFld) )
          aChecked   =  iif(lcase(aKeyVal) = lcase(trim(aCurVal)), " selected", "         ")
'  sayMsg("Option: " & aKeyVal & " = " & aValue)
          
      if (nWdt > 0) then  
'         aKeyVal    =    "'" & take( nWdt, replace( aValue, "'",  "''"  ) & "'" )
          aKeyVal    =   """" & take( nWdt, replace( aValue, """", """""") & """")
        else
          aKeyVal    =   """" &             replace( aValue, """", """""") & """"
          end if
          aOptions   =  aOptions & aFill & "  <OPTION value=" & aKeyVal & aChecked & ">" & aValue & "</OPTION>"
          pRS.MoveNext
          loop

          fmtOptions =  aOptions
      End Function

' --------------------------------------------------------------------------------

' %>