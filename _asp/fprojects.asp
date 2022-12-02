<%
Session("LastPage") = "_form.asp"
' ==========================================================================================
' SET CUSTOM VARIABLES
' ==========================================================================================

'    bNoEmails                =  true


'    bDebug                   =  true
    bDebug                   =  false


     bShowCustomFields        =  false

'    usesEmail                =  false
'    usesMenus                =  true
'    usesMarkedComplete       =  false

' -------------------------------------------------------------

    aNiceTableName           = "Projects"

    strTableName             = "tProjectList"
    aTableName               = "Member"
    strIDField               =  aTableName & "ID"
    strUpNav                 = ""

    strDefaultSort           =  Developer
    strDefDirection          = "ASC" ' or DESC

'    strFields4Add            = ""

    strFields4List           = "ProjectListID,Developer,Client,ProjectName,Dates,Industry"
'    strFields4Edit           = ""
'    strFields4Filter         =  strFields4Add

'    strCustomForm            = "form_app.htm"

'    Session("ShowRecords")   =  10

' -------------------------------------------------------------------------------------------

function  getFields4CustomForm( pFlds, pRS, bInp)

     dim  mFlds2()
   redim  mFlds2(         pFlds.count )

'         Set bShowCustomFields  = true  above to see field indices on form
'         mflds2( index on form) = setFld( TYPE, NAME, SIZE, pRS, RW)

       End Function

' -------------------------------------------------------------------------------------------

  Function SetCustomFields( pFld, pRS )

       End Function
' -------------------------------------------------------------------------------------------

  Function SetRequiredHiddenFields( pFld, pRS )

       End Function
' -------------------------------------------------------------------------------------------

  Function Check4Duplicates( Request )

       End Function
' -------------------------------------------------------------------------------------------

  Function Check4Validation( pRS)

       End Function
' -------------------------------------------------------------------------------------------

  Function SavCustomFields( pFld, byval aValue)

       End Function

%><!--#INCLUDE FILE="_formX.asp"    -->
