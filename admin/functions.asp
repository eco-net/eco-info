<%
Function OrgHas(t,orgid)
SET rsOH = Server.CreateObject("ADODB.recordset")
	rsOH.ActiveConnection=MM_ecoinfo_STRING
	rsOH.Source="SELECT * FROM " & t & " WHERE orgid=" & orgid
	rsOH.Open()
	if not rsOH.EOF then
	theid=rsOH("orgid")
	else
	theid=0
	end if
	rsOH.Close
	OrgHas=theid
end function
Function orgname(id)

name=""
set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.Source = "SELECT *  FROM eiorg_maindata  WHERE id=" & id
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3
rs.Open()
name=rs("name")
firstname=rs("firstname")
lastname=rs("lastname")
rs.Close
if Len(name)=0 then
name=firstname & " " & lastname
end if
orgname=name
end function

Function catsubname(cs,id)'returnerer navn ud fra id i enten kategori eller emne

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
if cs="cat" then
rs.Source = "SELECT name FROM eiorg_cat_maindata WHERE id=" & id
end if
if cs="calcat" then
rs.Source = "SELECT name FROM eical_cat_maindata WHERE id=" & id
end if
if cs="artcat" then
rs.Source = "SELECT name FROM eiart_cat_maindata WHERE id=" & id
end if
if cs="pubcat" then
rs.Source = "SELECT name FROM eilib_cat_maindata WHERE id=" & id
end if
if cs="sbj" then
rs.Source = "SELECT name FROM eisbj_maindata WHERE id=" & id
end if
if cs="opsl" then
rs.Source = "SELECT name FROM eiopsl_cat_maindata WHERE id=" & id
end if
rs.CursorType = 2
rs.CursorLocation = 3
rs.LockType = 2
rs.Open()
if not rs.EOF then
thename=rs("name")
else 
thename=""
end if
rs.Close
catsubname=thename
end function
%>