<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Server.ScriptTimeOut=1000
%>
<%
set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.Source = "SELECT firstname, middlename, lastname, fullname, id  FROM eiorg_maindata"
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3
rs.Open()
rs_numRows = 0
%>
<% 
while not rs.EOF
str=""
if rs("firstname")<>"" then
	str=rs("firstname")
end if
if rs("middlename")<>"" then
	str=str & " " & rs("middlename")
end if
if rs("lastname")<>"" then
	str=str & " " & rs("lastname")
end if
if str<>"" then
	rs("fullname")=str
	rs.Update
end if
rs.MoveNext
wend
%>
<%
rs.Close()
%>
