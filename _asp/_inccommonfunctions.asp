<% ' Common Function Library

' function getRS(           aSQL)
' function getMTRecordset(  aSQL)
' function displayField(    pFld,        aVal, byVal aHTML)
' function htmDate(         pFld,  byVal aVal, n)
' function htmFld(          pFld,  byVal aVal, byVal aJst)
' function fmtFldValue(     pFld)
' function fmtFValue(       pFld,  byVal aValue)
' function fmtPhone(  byval aStr)                                ' .(30625.09.1 BEG
' function setValidation(   aFld,  aTopMsg, aFldMsg)
' function numOnly(   byval aStr)
' function setValidation(   aFld,  aTopMsg, aFldMsg)
' function chkRequired(     aFld,  aName, aMsg)
' function chkValidation(   pRS )
' function clrValidation(   pRS )
' function bldFilter(       pRS )
' function getAllFields(    pRS )
' function getAllFieldsBut( pRS,   aFlds2Exclude )
' function inFlds(aList,    aFld)
' function resetLastValue(  aFld,  byVal aVal)
' function fmtSelect(       aFill, nWdt, aName aCurVal)
' function fmtSelect1(      aFill, nWdt, byval aKeyVal, aTable, aKeyFld, aValFld, aSrtFld, aWhere)
' function chkValDate(      aFld,  aName, aMsg)
' function chkOption(       aOpt,  aVal)


' function chkItm( aKeyVal, aKey,  aType)
' function chkBtn( aKeyVal, aKey)
' function chkBox( aKeyVal, aKey)
' function chkSel( aKeyVal, aKey)
' function fmtStr( aStr)
' function fmtNum( nNum)
' function fmtCur( nAmt)
' function fmtDte( dDate)
' function fmtTxt( aStr)

' function take( n,a)
' function cNum( n)
' function nMin( a,b)
' function nMax( a,b)
' function  iif( a,b,c)

' function msgIfError(       aFld)
' function Response_Redirect(aURL)

' --------------------------------------------------------------------------------

function getRS(aSQL, pFlds)
     Dim pRS ' Return a Collection for One Record, not 1st Recordset.  Preserves TEXT fields and holds ""s if Record not found
      on error resume next: if (cnum(bDeBug) = true) then on error goto 0

     Set pRS  = Conn.Execute(aSQL)

     if (Err.Number <> 0) then
         response.write "<font color=red><b>INVALID SQL STATEMENT: " & aSQL & "</b></font><br>"
         setValidation "","[3.2] SQL: "   & aSQL, ""
         setValidation "","ERROR: " & Err.Description, ""
'        response_redirect  3.2, RootFileName & "?ID=" & session("ID"), ""
         response.end
         end if

      on error goto 0

     Set getRS = Server.CreateObject("Scripting.Dictionary")

         bNewRec = (pRS.EOF or 0 = Session("ID"))
     For Each pFld in pRS.Fields
     if (bNewRec) then
         getRS.Add pFld.Name, ""
       else
         getRS.Add pFld.Name, pFld.Value
         end if
         Next

     set pFlds = pRS.Fields
         pRS.Close
    end function

 function getMTRecordset(aSQL)

          aSQL = left(aSQL, instr(ucase(aSQL) & " WHERE", " WHERE"))
          aSQL = replace(aSQL, "SELECT ", "SELECT TOP 1 ")

      Set pRS = Server.CreateObject("ADODB.Recordset")
          pRS.Open aSQL, Conn

      Set pRS  = Conn.Execute(aSQL)

'   response.write "FY: " & pRS("FY") & "<br>"
      For Each pField in pRS.Fields
          pField.value = null
          Next
'   response.write "FY: " & pRS("FY") & "<br>"

      Set getMTRecordset = pRS

      End Function
' --------------------------------------------------------------------------------

function displayField(pFld, aVal, byVal aHTML)

'        pFld.Name
'        pFld.Value
'        pFld.Type
'        pFld.DefinedSize
'        pFld.ActualSize

'   if (left(ucase(session("RO")),1) = "Y") then

'                aHTML = "<td align=""top"" ><font size=""3"">{FieldLabel}:&nbsp;</font></td>" & vbCrLf & "           "
'        aHTML = aHTML & "<td align=""top"" ><font size=""3"">{FieldValue}</font></td>"

'     else
'                aHTML = "<td align=""left"" valign=""top"" ><font size=""3"">{FieldLabel}</font></td>"
'        aHTML = aHTML & "<td align=""left"" valign=""top"" ><INPUT TYPE=""TEXT"" NAME=""{FieldName}"" VALUE=""{FieldValue}"" SIZE=""{FieldSize}""></td>"

'      end if

         aHTML = replace( aHTML, "{FieldLabel}", pFld.Name)
'        aHTML = replace( aHTML, "{FieldName}",  pFld.Name)
'        aHTML = replace( aHTML, "{FieldSize}",  nMin(pFld.DefinedSize + 5, 60))

'        aHTML = replace( aHTML, "{FieldValue}", aVal)
         aHTML = replace( aHTML, "{FieldValue}", htmFld(pFld, aVal, "left"))

         displayField  =  aHTML

     end function
' --------------------------------------------------------------------------------

 function htmlDate(pFld, byVal aVal, n)

      if (isDate(aVal)) then
          aVal = formatdatetime(aVal,n)
        else
'         aVal = ""
          end if
          htmlDate = htmlFld(aFld, aVal, "right", 10)
      end function
' --------------------------------------------------------------------------------

 function htmDate(pFld, byVal aVal, n)

      if (isDate(aVal)) then
          aVal = formatdatetime(aVal,n)
        else
'         aVal = ""
          end if
          htmDate = htmFld(pFld, aVal, "right")
      end function
' --------------------------------------------------------------------------------

 function htmlFld(aFld, byVal aVal, byVal aJst, nLen)

          aJst  = ucase(aJst)

      if (left(ucase(session("RO")),1) = "Y") then
          aVal  = "&nbsp;" & aVal & "&nbsp"
        else
          aVal  = "<INPUT NAME=""" & aFld & """ VALUE=""" & aVal & """ SIZE=""" & nLen & """>"
          aJst = "LEFT"
          end if

          htmlFld = "<SPAN CLASS=""InpBox"" ID=""" & aFld & """ ALIGN=""" & aJst & """>" & aVal & "</SPAN>"
      end function

 function htmFld(pFld, byVal aVal, byVal aJst)

          aFld  = pFld.Name
          nLen  = nMin(pFld.DefinedSize + 5, 60)
          aJst  = ucase(aJst)

      if (left(ucase(session("RO")),1) = "Y") then
          aVal  = "&nbsp;" & aVal & "&nbsp"
        else
          aVal  = "<INPUT NAME=""" & aFld & """ VALUE=""" & aVal & """ SIZE=""" & nLen & """>"
          aJst = "LEFT"
          end if

          htmFld = "<SPAN CLASS=""InpBox"" ID=""" & aFld & """ ALIGN=""" & aJst & """>" & aVal & "</SPAN>"

      end function
' -------------------------------------------------------------------------------- '

 function fmtFldValue( pFld)
          fmtFldValue = fmtSQL_FldEqRequestVal( pFld)
      end function

 function fmtSQL_FldEqRequestVal( pFld)
      dim aValue, aField

          aField =  pFld.name
          aValue =  Request(aField)

'     if (aField = "CreatedBy") then aValue = session("PersonEmail")    '*.(40619.04.1
'     if (aField = "CreatedAt") then aValue = now()                     '*.(40619.04.2

          fmtSQL_FldEqRequestVal = "[" & aField & "] = " & fmtSQL_Val( pFld, aValue)
      end function

 function fmtSQL_Val( pFld, aValue)
      dim aField

          aField =  pFld.name

'         aExpr  = "Replace(Ucase(""" & aValue & """),""'"",""''"")"
'         aValue =  eval(aExpr)
          aValue =  replace(aValue, "'", "''")

          nType  =  pFld.type
      if (nType  >  10 and not (nType = 135 or nType = 131) ) then

      if (nType  =  200) then  ' varchar

      if (0 < instr(aField, "Phone")) then aValue = fmtPhone(aValue)   ' .(30625.09.2 '

'         response.write aField & "[ " & pFld.definedSize & "]: '" & aValue & "'" & "<br>"
      if (len(aValue) > pFld.definedSize) then
          aValue = left( aValue, pFld.definedSize)
          setValidation pFld.Name, pFld.Name & " truncated to " & pFld.definedSize & " characters.", ""
          end if

          end if ' (nType = 200)

          aValue = "'" & trim(aValue) & "'"

        else     ' (nType <= 10)  Numeric

          if ("" < aValue & "") then
'             response.write pFld.Name & " is " & nType & " = " & aValue & "<br>"

              if (nType = 7 or nType = 135) then
                  aValue = "'" & formatdatetime(aValue) & "'"
                else
                  aValue = cnum(aValue)
                 end if
            else
              aValue = "NULL"
              end if

          end if

'         fmtFldValue =  aField & " = " & aValue
          fmtSQL_Val  =  aValue

      end function
' --------------------------------------------------------------------------------

 function fmtSQL_Flds( byVal aFlds )
          fmtSQL_Flds = aFlds
      if (aFlds = "*") then exit function
          aFlds = replace( aFlds, "," ,  ", " )
          aFlds = replace( aFlds, ", ", "], [")
          fmtSQL_Flds  = "[" & aFlds & "]"
      end function
' --------------------------------------------------------------------------------

 function fmtFValue(pFld, byVal aValue)

'         aExpr =  "Replace(Ucase(""" & aVal & """),""'"",""''"")"
'         aValue =  eval(aExpr)
          aValue =  replace(aValue & "", "'", "''")

'         response.write pFld.name & " (type = " & pFld.Type & ") is " & pFld.definedSize & "<br>"

          nType  =  pFld.type
      if (nType  >  10 and not (nType = 135 or nType = 131) ) then
      if (nType  = 200) then
      if (len(aValue) > pFld.definedSize) then

'         response.write "Changing width to " & pFld.definedSize & " from " & len(aValue) & "<br>"
          setValidation pFld.Name, pFld.Name & " truncated to " & pFld.definedSize & " characters.", ""

          aValue = left( aValue, pFld.definedSize)
          end if
          end if
          aValue = "'" & trim(aValue) & "'"
        else

         if ("" < aValue & "") then
'        response.write pFld.Name & " = " & aType & "<br>"
         if (nType = 7 or nType = 135) then
             aValue = "'" & formatdatetime(aValue) & "'"
           else
             aValue = cnum(aValue)
             end if
           else
             aValue = "NULL"
          end if
          end if

          fmtFValue =  aValue
     end function
' --------------------------------------------------------------------------------

 function fmtPhone(byval aStr)                                ' .(30625.09.1 BEG

            fmtPhone = aStr
        if ("" = aStr & "") then exit function

            aExt = "": nPos = instr(lcase(aStr), "x")
        if (nPos > 0) then
            aExt =  mid(aStr, nPos+1)
            aStr = left(aStr, nPos-1)
            end if

            aStr = numOnly(aStr): aExt = mid(aStr,11) & aExt
        if (7 >= len(aStr) and 0 = instr(aStr, "202")) then aStr = "202" & aStr
            aRts = "(" & left(aStr,3) + ") " & mid(aStr,4,3) + "-" + mid(aStr,7,4)
        if (aExt > "")     then aRts = aRts & " x" & left(numOnly(aExt),4)
            fmtPhone = aRts
       end function
' --------------------------------------------------------------------------------

 function numOnly(byval aStr)
        dim i
        for i = 1 to len(aStr)
            aChr = mid(aStr,i,1)
        if (0 < instr("0123456789", aChr)) then aRts = aRts & aChr
            next
            numOnly = aRts
        end function                                           ' .(30625.09.1 END
' --------------------------------------------------------------------------------

 function setValidation( aFld, aTopMsg, aFldMsg)

            Session("valCount")     = cnum(Session("valCount")) + 1
            Session("valMessage")   = Session("valMessage") & aTopMsg & "<br>"
        if (aFld > "") then
            Session("val" & aFld)   = aFldMsg
            end if
            setValidation           = true

    if (Session("valOldValues") = "") then
        for each aFld in Request.form
            Session("valOldValues") = Session("valOldValues") & "|" & aFld & "=" & Request.form(aFld)
            next
        end if

        end function
' --------------------------------------------------------------------------------

 function chkRequired( aFld, aName, aMsg)

'           response.write aFld & " is '" & request(aFld) & "'<br>"

        If ("" = trim( request(aFld) & "")) Then
            setValidation aFld, aName, aMsg
            chkRequired = true
          else
'           Session("val" & aFld) = ""
            chkRequired          = false
            End If

        end function
' --------------------------------------------------------------------------------

 function chkValidation( pRS )
        Dim aSQL

           for each pFld in pRS.Fields

'          If (0 = instr(1, strIDField & "," & strFields4Sys, pFld.Name, 1)) then
           If ( not inFlds( strIDField & "," & strFields4Sys, pFld.Name)) then

'                  response.write pFld.Name & " is " & pFld.Type & "<br>"
           if (not pFld.Attributes and &H00000040) then

'              if (ucase(pFld.Name) <> ucase(strTableName & "ID")) then
               if ( not inFlds( strIDField, pFld.Name)) then

                   chkRequired   pFld.Name, pFld.Name, pFld.Name & " is a required field."
                   end if
               end if
           if ("" < request(pFld.Name) & "") then
''                 response.write pFld.Name & " is " & pFld.Type & " = " & request(pFld.Name) & "<br>"

               if ((pFld.Type =  7 or pFld.Type = 135) and not(isDate(request(pFld.Name)))) then
                   setValidation pFld.Name, pFld.Name, Request(pFld.Name) & " is not a valid date."
                 else

               if ((pFld.Type < 10 or pFld.Type = 131) and not(isNumeric(request(pFld.Name)))) then
                   setValidation pFld.Name, pFld.Name, Request(pFld.Name) & " is not a valid number."
                   end if
                   end if
               end if
               end if
           next

        end function
' --------------------------------------------------------------------------------

   function clrValidation( pRS)

'           response.write "Clearing Validation<br>"

            aOldVals = Session("valOldValues")
         do while aOldVals > ""
            aFld     =  left(aOldVals, instr(aOldVals & "=", "=") - 1)
            aOldVals =   mid(aOldVals, instr(aOldVals & "|", "|") + 1)
            Session("val" & trim(aFld)) = ""
            loop

        if (isObject(pRS)) then
        for i = 0 to pRS.fields.count - 1
            Session("val" & pRS.fields(i).name) = Null
            next
        end if

            Session("valCount")         =  0
            Session("valMessage")       = ""
            Session("valOldValues")     = ""
            Session("valMessageBanner") = ""

        end function
' --------------------------------------------------------------------------------

 function bldFilter( pRS )
      Dim aStr, pFld, aTyp, aVal:  aStr = ""

      For Each pFld in pRS.Fields

'         if "" <> "" & Request(pFld.Name) then
'           aStr = aStr & pFld.Name & " Like '" & Request(pFld.Name) & "%' AND "
'         end if
'     if (not pFld.Attributes and &H00000040) then    ' Required

               aTyp = ""
               aVal = trim(request(pFld.Name) & "")
           if (0 < instr("<>=", left(aVal,1))) then
               aTyp = left( aVal, iif("=" = mid(aVal,2,1), 2, 1))
               aVal = trim( mid( aVal, 1 + len(aTyp)))
               end if
      if ("" < aVal) then

          if (pFld.Type =  7 or pFld.Type = 135) then ' Dates

                  aTyp = iif(aTyp > "", aTyp, "Between")
              if (not isDate( aVal) ) then
                  setValidation pFld.Name, pFld.Name, Request(pFld.Name) & " is not a valid date."
                else
                  aStr = aStr & pFld.Name & " " & aTyp & " '" & fmtDate( int(cdate(aVal))) & "' AND '"
              if (aTyp = "Between") then
                  aStr = aStr & fmtDate( dateadd("d", 1, int(cdate(aVal)))) & "' AND "
                  end if
                  end if

      elseif (pFld.Type < 10 or pFld.Type = 131) then ' Numbers

                  aTyp = iif(aTyp > "", aTyp, "=")
                  aVal = replace( replace( aVal, "$",""), ",", "")
              if (not isNumeric(aVal) ) then
                  setValidation pFld.Name, pFld.Name, Request(pFld.Name) & " is not a valid number."
                else
                  aStr = aStr & pFld.Name & " " & aTyp & " " & cnum(aVal) & " AND "
                  end if
            else                                      ' Characters

                  aTyp = iif(aTyp > "", aTyp, "Like")
                  aStr = aStr & pFld.Name & " " & aTyp & " '" & aVal & iif("Like" = aTyp,"%","") & "' AND "

              end if
          end if' ("" < aVal)

          Next

'         Session("Where4List") = left( aStr,         len( aStr )-5)
          bldFilter             = left( aStr, nMax(0, len( aStr )-5))
      End Function
' --------------------------------------------------------------------------------

 function getAllFields( pRS)
      Dim aStr, pFld:  aStr = ""
      For Each pFld in pRS.Fields
          aStr = aStr & "," & pFld.Name
          Next
          getAllFields = mid(aStr,2)
      End Function
' --------------------------------------------------------------------------------

 function getAllFieldsBut( pRS, aFlds2Exclude )
      Dim aStr,pFld: aStr = ""
      For Each pFld in pRS.Fields
      if (not inFlds( aFlds2Exclude, pFld.Name)) then
          aStr = aStr & "," & pFld.Name
          end if
          Next
          getAllFieldsBut = mid(aStr,2)
      End Function
' --------------------------------------------------------------------------------

 function inFlds(byval aList, aFld)
          aList  = replace(aList, " ", ""):                                         ' .(40619.04.1
          inFlds = (0 < instr( 1, "," & aList & ",", "," & trim(aFld) & ",", 1))
'         response.write "inFlds('" & aList & "', '" & aFld & "')<br>"
'         response.write "inFlds(strADDfields, '" & aFld & "') = " & inFlds & "<br>"
      End Function
' --------------------------------------------------------------------------------

 function resetLastValue(aFld, byVal aVal)

          nPos = instr(session("valOldValues"), aFld & "=")
      if (nPos > 0) then
          aVal = mid(session("valOldValues"), nPos + len(aFld) + 1)
          aVal = trim(left(aVal, instr(aVal & "|", "|") - 1))
'         response.write "resetLastValue for " & aFld & ": " & aVal & "<br>"
          end if
          resetLastValue = aVal
      End Function

' --------------------------------------------------------------------------------

 function fmtSelect( aFill, aName, aCurVal)
      Dim nWdt, aTable, aKeyFld, aValFld, aWhere, aSQL, pRS1

          nWdt       =  40            ' Width of <OPTION Value="{%= take( aKeyVal, nWdt) %}">{%= aValue %}</OPTION>
          aFill      =  vbCrLf & ""   ' spaces before <SELECT tag

          aTable     = "tblLkUp"
          aKeyFld    = "LkUpValue"
          aValFld    = "LkUpValue"
          aWhere     = "WHERE (LkUpType = '" & aName & "')"

      if (aKeyFld   <>  aValFld) then
          aKeyFld    =  aKeyFld & "," & aValFld
          end if

          aSQL       = "SELECT " & aKeyFld & " FROM " & aTable & " " & aWhere & " ORDER BY " & aValFld

      Set pRS1       =  Conn.Execute( aSQL)

      dim aStr:        aStr   = aFill & "  <td>" & aName & "</td>"
          aStr      =  aStr   & aFill & "  <td>"
          aStr      =  aStr   & aFill & "    <select name=""" & aName & """>"
          aStr      =  aStr   & aFill & "      <option value=""""></option>"
          aStr      =  aStr   &                  fmtOptions( aFill, aName, aCurVal, pRS1, nWdt )
          aStr      =  aStr   & aFill & "    </select>"
          aStr      =  aStr   & aFill & "  </td>"

          pRS1.close
      Set pRS1       =  Nothing

          fmtSelect = aStr
      end function

 function fmtOptions( aFill, aName, aCurVal)
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

 function fmtSelect1( aFill, nWdt, byval aKeyVal, aTable, aKeyFld, aValFld, aSrtFld, aWhere)

      dim aStr, WHERE, ORDER_BY, aSQL, aValue, aChecked, pRS
          aStr      = "":     aKeyVal = ucase(aKeyVal)
          WHERE     = "": if (aWhere  > "") then WHERE    = " WHERE "      & aWhere
          ORDER_BY  = "": if (aSrtFld > "") then ORDER_BY = " ORDER BY ["  & aSrtFld & "]"
          aSQL      = "SELECT [" & aKeyFld & "], [" & aValFld & "] FROM [" & aTable  & "]" & WHERE & ORDER_BY
'         response.write "SQL: " & aSQL    & "<br>" & vbCrLf
'         response.write "Val: " & aKeyVal & "<br>" & vbCrLf
'         response.end
      set pRS       = Conn.Execute(aSQL)
       do while not pRS.EOF
          aValue    = ucase(trim(pRS( aKeyFld))): aValueQ = replace( aValue, """", """""") & """"
          if (aKeyVal = aValue) then  aValueQ  =  aValueQ & " SELECTED"
          aStr      = aStr & vbCrLf & aFill & "<option value=""" & take( nWdt + 10, aValueQ) & ">" & take(60, pRS( aValFld)) & "</option>"
''        response.write "'" & aKeyVal & "' = '" & aValue & "'" & vbCrLf
          pRS.MoveNext
          loop
          pRS.Close
      set pRS = Nothing
          fmtSelect = mid( aStr, 1 + len(aFill))
      end function
' --------------------------------------------------------------------------------

 function chkValDate(aFld, aName, aMsg)
      if (not(isDate(request(aFld))) and request(aFld) > "") then setValidation aFld, aName, aMsg
      end function
' --------------------------------------------------------------------------------

 function chkOption( aOpt, aVal)
'         response.write aOpt & " = " & aVal & " = " & (aOpt = aVal) & "<br>"
          chkOption = "'" & replace(aOpt,"'","''") & "' " & iif(aOpt = aVal, "selected", "")
      end function
' --------------------------------------------------------------------------------

 function chkItm( aKeyVal, aKey, aType)
      dim aStr:    aStr = """" & aKey & """ "
      if (ucase(aKeyVal) = ucase(aKey)) then
          chkItm = aStr & aType & "ED"
        else
          chkItm = aStr & take( len(aType) + 2, "")
          end if
      end function
' --------------------------------------------------------------------------------

 function chkBtn( aKeyVal, aKey)
          chkBtn = chkItm( aKeyVal, aKey, "CHECK")
      end function
' --------------------------------------------------------------------------------

 function chkBox( aKeyVal, aKey)
          chkBox = chkItm( aKeyVal, aKey, "CHECK")
      end function
' --------------------------------------------------------------------------------

 function chkSel( aKeyVal, aKey)
          chkSel = chkItm( aKeyVal, aKey, "SELECT")
      end function
' --------------------------------------------------------------------------------

 function fmtStr( aStr)
          fmtStr = replace( aStr & "", chr(34), "&quot;")
      end function
' --------------------------------------------------------------------------------

 function fmtNum( nNum)
      if (not isNumeric( nNum)) then
          fmtNum = ""
        else
          fmtNum = formatNumber( cdbl(nNum), 0)
          end if
      end function
' --------------------------------------------------------------------------------

 function fmtCur( nAmt)
      if (not isNumeric( nAmt)) then
          fmtCur = ""
        else
          fmtCur =  formatCurrency( cdbl(nAmt & ""), 1)
          end if
      end function
' --------------------------------------------------------------------------------

 function fmtDate( dDate)
          fmtDate = datepart("m",dDate) & "/" & datepart("d",dDate) & "/" & datepart("yyyy",dDate)
      if (dDate = int(dDate)) then exit function
          fmtDate = fmtDate & " " & datepart("h",dDate) & ":" & datepart("n",dDate) & ":" & datepart("s",dDate)
     end function
' --------------------------------------------------------------------------------

 function fmtDte( dDate)
          fmtDte = dDate
      end function
' --------------------------------------------------------------------------------

 function fmtTxt(aStr)
          fmtTxt = aStr
      end function
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

 function msgIfError( aFld)

          aErrorMsg = Session("val" & aFld) & ""
      if (aErrorMsg > "" AND Session("valCount") > 0) then
          msgIfError = "        <TR><TD></TD><TD><font size=""3"" color=""red""><b> * " & aErrorMsg & "</b></font></TD></TR>" '  & vbCrLf
        else
          msgIfError = ""
          end if
      end function
' --------------------------------------------------------------------------------

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

      response.redirect aPage & aProcess & iif(bDebug,"&debug","")

  End Function
' --------------------------------------------------------------------------------

    %>