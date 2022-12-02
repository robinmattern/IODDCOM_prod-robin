<%@ Language=VBScript %>
<% Response.Buffer = True %>
<!--#INCLUDE FILE="_incexpires.asp"-->
<!--#INCLUDE FILE="_inccreateconnection.asp"-->
<!--#INCLUDE FILE="_incheader.asp"-->
<%


'HYPERLINKS FOR "GENERATE FILE"
'---------------------------
Session("GenerateCsvHyperlink")="_generatefile.asp?type=csv"
Session("GenerateXmlHyperlink")="_generatefile.asp?type=xml"
Session("GenerateASCIIHyperlink")="_generatefile.asp?type=comma"

'Color Control Variables
'-----------------------
UsingArrays = "Yes"
If Len(Trim(Session("ReportVariablesAreSet"))) = 0 then
	UsingArrays = ""
' REPORT TITLE LINE
'The Background for TITLE line with the report name is BLUE 
Session("ReportHdBCK")="#000080"	
'The text for line with the report name is WHITE
Session("ReportHdTXT")="#FFFFFF"

' REPORT HEADER LINE
' Top - header line of the report had BLUE backgroud
Session("TopRowBCK")="#000080"
' and WHITE text on it
Session("TopRowTXT")="#FFFFFF"

' ODD LINES
' The each next Odd line has Light Gray Background
Session("OddLineBCK")="#CAEEFF" '"#99CCFF"
Session("OddLineTXT")="#000000"

' EVEN LINES
' The each Even Line is Light blue
Session("EvenLineBCK")="#FFFFFF"
Session("EvenLineTXT")="#000000"

'TEXT SIZE VARIABLES
'-------------------
Session("TitleHeight")=20
Session("TitleNameFontSize")=3
Session("TitleDateFontSize")=2
Session("HeaderFontSize")=2
Session("LinesFontSize")=2
Session("SelectStatusFontSize")=2
Session("SaveReportToFileFontSize")=2

'COLUMNS SIZE AND WRAP CONTROLL
'-------------------
NumberOfColumns=10

End If
Session("ReportVariablesAreSet") = ""

If len(trim(Session("SQLStrReport"))) = 0 then
	Session("SQLStrReport") = Session("SQlStr")
	If len(trim(Session("SQLStrReport"))) = 0 then
		Session("Message") = "Report error: SqlStrReport is empty."
		response.redirect "_message.asp"
	End If
End if

'Get row count
From = mid(Session("SQLStrReport"),Instr(UCASE(Session("SQLStrReport"))," FROM "))
If instr(From, "ORDER BY") > 0 then
	From = LEFT(From,Instr(UCASE(From),"ORDER BY")-1)
End if	
'response.write "SQLStrReport: " & Session("SQLStrReport") & "<br>"
'response.write "FROM: " & From & "<br>"
'response.end
SQLStrCount = "SELECT COUNT(*) AS Rows " & From
'response.write "SQLStrCount: " & SQLStrCount & "<br>"
'response.end

Set rs = Conn.Execute(SQLStrCount)
RowCount = int(rs("rows"))
If RowCount = 0 then
	Session("Message") = "No records found.<br><br>" & Session("SQLStrWhere") & "."
	response.redirect "_message.asp"
End if	
'response.write "RowCount: " & RowCount & "<br>"

Set rs = Conn.Execute(Session("SQLStrReport"))

FormattedDate =Right(FormatDateTime(date(),1),len(FormatDateTime(date(),1))-instr(FormatDateTime(date(),1),",")) & " " & time() '& " GMT"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">

<HTML>
<HEAD>
<TITLE>Report</TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 6.0">

</HEAD>

<BODY LINK="#000080" VLINK="#800080" TEXT="#000000">
<%
dim ReportHdBCK, ReportHdTXT, TopRowBCK, TopRowTXT, OddLineBCK, EvenLineBCK, ActualBCK, ForcedCut
dim SeparateRows

'The backgroud for line with the report name 
ReportHdBCK=Session("ReportHdBCK")
'The text for line with the report name 
ReportHdTXT=Session("ReportHdTXT")
' Top - header line of the report 
TopRowBCK=Session("TopRowBCK")
' and YELLOW text on it
TopRowTXT=Session("TopRowTXT")
' The each Odd line 
OddLineBCK=Session("OddLineBCK")
OddLineTXT=Session("OddLineTXT")
' The each Even Line 
EvenLineBCK=Session("EvenLineBCK")
EvenLineTXT=Session("EvenLineTXT")

' If there is no possibility to insert space for wraping, the string for
' displaying in browser will be cutted each ForcedCut chars. Recommended default is 12
VarForcedCut=15


'********************************************************************************
FUNCTION StringToWrap (myString, MaxUnwrapChars)
'********************************************************************************
' Returns substring, which has no space inside and is longer than MaxUwrapChars 
' If there is more such substrings, it returns FIRST from the RIGHT side, since
' this substring is usually causing "wrapping problems"
'
' How to use example:
' ActualStringToWrap=StringToWrap("1iuoiuo2 3werwerwerer4 5oiuoiuoiuouoiuoiuoiuoiu6 7oiuouio8",12)
' IF LEN(ActualStringToWrap)>0 THEN
' 	response.write "We will wrap:"
'	response.write ActualStringToWrap
' END IF
'

	StringToWrap=""
	if right(mystring,1)<>" " then
		mystring=mystring+" "
	end if

	DO WHILE LEN(mystring)>0 and instr(mystring," ")>0

		cutstring=left(mystring,instr(mystring," "))
		mystring=mid(mystring, instr(mystring," ")+1,len(mystring)-instr(mystring," "))
		IF len(cutstring)>MaxUnwrapChars then
			StringToWrap=cutstring
		END IF
	LOOP
	
END FUNCTION

'*****************************************************************************
FUNCTION AddWrapSpace(mystring,myMaxStringSize)
' THIS IS MAIN "WRAP" FUNCTION
'*****************************************************************************

if IsNull(mystring) then
	mystring = "&nbsp;"
end if



' All type of variables are converted to the string type - it allows wraping even for 
' DATE


mystring = Cstr(mystring)

outstring=""

if right(mystring,1)<>"," then
mystring=mystring+","
end if

DO WHILE LEN(mystring)>0 and instr(mystring,",")>0

	outstring=outstring+left(mystring,instr(mystring,","))
	mystring=mid(mystring, instr(mystring,",")+1,len(mystring)-instr(mystring,","))

	if len(trim(left(mystring,1)))=0 then
		'The space is behind coma, we do not need space for wrapping
	else
		mystring=" "+mystring
	end if
LOOP
outstring=LEFT(outstring, len(outstring)-1)

' We will look, whether there is some part of the string biger than out limit
' and whether contains char -
tempstring=outstring
ActualStringToWrap=StringToWrap(tempstring,myMaxStringSize)
IF LEN(ActualStringToWrap)>0 AND instr(outstring,"-") THEN
	outstring=AddWrapSpace2(outstring,"-")
END IF


' We will look, whether there is some part of the string biger than out limit
tempstring=outstring
ActualStringToWrap=StringToWrap(tempstring,myMaxStringSize)
IF LEN(ActualStringToWrap)>0 AND instr(outstring,"/") THEN
	outstring=AddWrapSpace2(outstring,"/")
END IF

'Last Check is for ONE LONG STRING with still no spaces !!!
tempstring=outstring
tempstring2=outstring
ActualStringToWrap=StringToWrap(tempstring,myMaxStringSize)
IF LEN(ActualStringToWrap)>0 AND INSTR(TRIM(tempstring2)," ")=0 THEN
	Outstring=HardForcedCut(OutString,myMaxStringSize)
END IF

if trim(outstring)="" then
	outstring = "&nbsp;"
end if

AddWrapSpace=outstring
END Function

'***************************************************************************
FUNCTION AddWrapSpace2(mystring,wrapchar)
'***************************************************************************

outstring=""

if right(mystring,1)<>wrapchar then
mystring=mystring+wrapchar
end if

DO WHILE LEN(mystring)>0 and instr(mystring,wrapchar)>0

	outstring=outstring+left(mystring,instr(mystring,wrapchar))
	mystring=mid(mystring, instr(mystring,wrapchar)+1,len(mystring)-instr(mystring,wrapchar))

	if len(trim(left(mystring,1)))=0 then
		'The space is behind WARPCHAR, we do not need space for wrapping
	else
		mystring=" "+mystring
	end if
LOOP
outstring=LEFT(outstring, len(outstring)-1)
AddWrapSpace2=outstring
END Function


Function HardForcedCut(mystring,CutValue)
	tempstring=mystring+" "
	outstring=""
	DO WHILE len(tempstring)>CutValue
		outstring=outstring+left(tempstring,CutValue)+" "
		tempstring=mid(tempstring,CutValue+1,len(tempstring)-len(left(tempstring,CutValue)))
	
	LOOP	
	
	outstring=outstring+" "+tempstring
	
	HardForcedCut=outstring
END Function


%>

  <center>

<div align="left">

<TABLE  WIDTH=640 BORDER=0 CELLSPACING=0 CELLPADDING=0 style="border-collapse: collapse" bordercolor="#111111">
<TR VALIGN="top" ALIGN="left">
	<TD HEIGHT =<%=Session("TitleHeight")%> bgcolor=<%=ReportHdBCK%> ALIGN="Left" valign="bottom"><b><font color=<%=ReportHdTXT%> size=<%=Session("TitleNameFontSize")%>>&nbsp;<%=Session("ReportName")%>&nbsp;&nbsp;<small>(&nbsp;<%=RowCount%>&nbsp; rows returned)</small></font></b></td>
	<TD HEIGHT =<%=Session("TitleHeight")%> bgcolor=<%=ReportHdBCK%> ALIGN="Right" valign="bottom"><b><font color=<%=ReportHdTXT%> size=<%=Session("TitleDateFontSize")%>><%=FormattedDate%>&nbsp;</font></b></TD>
</TR>
<TR >
	<TD ALIGN="Left" valign="bottom"><font size=<%=Session("SelectStatusFontSize")%> color=<%=OddLineTXT%>><%=Session("SQLStrWhere")%></font></TD>
<%'Session("SQLStrWhere")= ""%>
	<TD ALIGN="Right" valign="bottom"><font size=<%=Session("SaveReportToFileFontSize")%> color=<%=OddLineTXT%> >
	<a title="Creates a, Comma Separated Value file (.csv) for use in spreadsheet programs." href="<%=Session("GenerateCsvHyperlink")%>">Spreadsheet</a>
	<a title="Creates an Extended Markup Language (.xml) file from this data." href="<%=Session("GenerateXmlHyperlink")%>">XML</a>&nbsp;<a title="Creates an American Standard Code for Information Interchange file from this data." href="<%=Session("GenerateASCIIHyperlink")%>"> ASCII</a>&nbsp;</font></TD>

</TR>
</Table>
</div>
<%
	' This is background color for first displayed line under Header Line
	ActualBCK=EvenLineBCK
	
 %>
<div align="left">
	 <TABLE  WIDTH=640 BORDER=1 CELLSPACING=0 CELLPADDING=2 style="border-collapse: collapse" bordercolor="#111111" >
	 <TR VALIGN="top" ALIGN="left">
<%
	For Each MyField in rs.Fields
		If Ucase(Right(Trim(MyField.Name),2)) <> "ID" then 
%>
			<TD valign="bottom" bgcolor=<%=TopRowBCK%>><P><B><font size=<%=Session("HeaderFontSize")%> color=<%=TopRowTXT%>><% Response.write MyField.Name %></font></B></TD>

<%		End If
	Next
%>
	 </TR>

<%
If UsingArrays = "Yes" then

' The Array with the arrColumnPixelWidth is stored in session variable "StorearrColumnPixelWidth"
' we have to build local array first

	arrColumnPixelWidth = Session ("arrStoreColumnPixelWidth")

' The next Array contains the WRAP sizes for columns
	arrColumnWrapSize=Session("arrStoreColumnWrapSize")

End If	

	Do Until rs.EOF
	
		For i = 0 to rs.Fields.Count - 1
			FieldVal = rs.Fields(i)
			'Response.write fieldval
			If isnull(fieldVal) or fieldval = "" or fieldVal = " " then
				FieldVal = "&nbsp;"
			end If 
			If Ucase(Right(Trim(rs.fields(i).Name),2)) <> "ID" then 

				If UsingArrays = "Yes" then %>
		  			<TD width=arrColumnPixelWidth(i,0) <%=arrColumnPixelWidth(i,1)%> bgcolor= <%=ActualBCK%> ><font size="2"><% response.write AddWrapSpace(FieldVal,arrColumnWrapSize(i)) %></font>&nbsp;</TD>
				<% Else %>
		  			<TD bgcolor= <%=ActualBCK%> ><font size="2"><% response.write FieldVal %></font>&nbsp;</TD>
				<% End If %>
	 	<%	Else %>
	  			<!--TD bgcolor= <%=ActualBCK%> ><font size="2"><% response.write FieldVal %></font>&nbsp;</TD-->
		<%	End if
		Next
		
		if ActualBck=OddLineBCK  then 	
				ActualBck=EvenLineBCK
		Else
				ActualBck=OddLineBCK
		End if
		
%>
<%
rs.MoveNext%>

		</TR>

<%Loop%>
  
  </center>
<p>
<br>
</TABLE>
</div>
</BODY>