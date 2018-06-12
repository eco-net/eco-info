<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%if request("stem")<>"" then
Dim Conn, rs
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.Recordset")
Conn.Open MM_ecoinfo_STRING

i=1
For each item in request("checkbox")
	rs.Open "eiafstem_ning", Conn, 1, 3, 2
	while not rs.EOF
		if CStr(rs("id"))=CStr(item) then
		'response.write item
		'response.end
		stemmer=rs("stemmer")
		rs("stemmer")=stemmer + 1
		rs.Update
		end if
	rs.movenext
	wend
	rs.Close
next
Conn.Close
if request("mail2")<>"din@mailadresse" and request("mail2")<>"" then
theSQL="INSERT INTO eiafstem_konkurrence (emne_id,mail) VALUES (" &_
	request("afsid2") & "," &_
	"'" & replace(request("mail2"),"'","''") & "'" &_
	")"
	execCommand theSQL
end if
end if

%>
<script language="javascript">
alert("Tak for deltagelsen");
window.close();
</script>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">

</body>
</html>
