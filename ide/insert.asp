<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="shared/common.asp" -->
<%
tom="empty"

theSql="INSERT INTO eiide_maindata (name,email,title,shortdescr,descr) " &_
"VALUES ('" & request("name") & "','" &_
request("email") & "','" &_
request("title") & "','" &_
request("shortdescr") & "','" &_
tom & "')"


execCommand theSQL

response.redirect("index.asp")


%>
