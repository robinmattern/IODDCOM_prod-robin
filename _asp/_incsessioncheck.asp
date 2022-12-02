<%
if pagetitle <> "Score 300" then
	if trim(session("sessionid"))="" then
		response.redirect  "default.asp"
	end if	
end if
If len(trim(session("changedby"))) = 0 then
	response.redirect "default.asp"
End if 

%>