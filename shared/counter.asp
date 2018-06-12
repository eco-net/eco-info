<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Function count()
set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.Source = "SELECT *  FROM stat_counter  ORDER BY date DESC"
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 2
rs.Open()

if rs("date")< date then
response.write(date)
rs.AddNew
rs("date")=date
rs("ecoinfo") = 1
rs.Update
else
count=rs("ecoinfo") + 1
rs("ecoinfo")=count
rs.Update
end if
rs.Close()
end Function
%>
