<%
remote=Request.ServerVariables("REMOTE_ADDR")
'response.write remote
'response.end

'IF remote<>"80.160.41.182" THEN response.redirect "http://www.eco-info.dk"
IF remote<>"82.180.34.145" THEN response.redirect "http://www.eco-info.dk"
%>