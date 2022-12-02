<% 
response.expires = 0
response.expiresabsolute = Now() - 1
response.addHeader "pragma","no-cache"
response.addHeader "cache-control","private"
response.CacheControl = "no-cache"
%>
