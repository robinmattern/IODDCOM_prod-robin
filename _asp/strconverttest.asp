<%@ LANGUAGE="VBScript" %> <%Response.Buffer = True%> 

<!--#INCLUDE FILE="_incstrconvert.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
	<%
	response.write strconvert("robin.mattern@inetpurchasing.com","html") & "<BR>"
	response.write strconvert("robin.mattern@inetpurchasing.com","ascii") & "<BR>"
	response.write strconvert("robin.mattern@inetpurchasing.com","password") & "<BR>"
	response.write strconvert("ab","ascii") & "<BR>"
	response.write strconvert("ab","password") & "<BR>"
	%>
</body>



</html>