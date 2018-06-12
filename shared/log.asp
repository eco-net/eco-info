<!--#include virtual="/connections/ecoinfo.asp" -->

<%

remote=Request.ServerVariables("REMOTE_ADDR")
if not Session("userid")=Session.SessionID then
	if remote<>"80.160.41.182" then
		addToAllDays("users")
	else
		addToAllDays("eusers")
	end if
Session("userid")=Session.SessionID
end if

Function reg(page)
remote=Request.ServerVariables("REMOTE_ADDR")
if remote<>"80.160.41.182" then
	addToAllDays(page)
else
	addToAllDays("e" & page)
end if
end Function


Function addToAllDays(att)

set rsAll = Server.CreateObject("ADODB.Recordset")
rsAll.ActiveConnection = MM_ecoinfo_STRING
rsAll.Source = "SELECT * FROM eisys_ecoinfo_count"
rsAll.CursorType = 0
rsAll.CursorLocation = 2
rsAll.LockType = 2
rsAll.Open()
if rsAll(att)=0 then
		c=1
	else
		c=rsAll(att) + 1
	end if
	rsAll(att)=c
	rsAll.Update
rsAll.Close
end Function





%>