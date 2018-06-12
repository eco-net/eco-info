<!--#include virtual="/Connections/ecoinfo.asp" -->
<%Dim ID
ID=request("ID")
set conn=Server.CreateObject("ADODB.Connection")
set rs= Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM bu_Artikel where Artikel_ID= " & ID 
conn.open MM_ecoinfo_STRING
rs.Open strSQL, conn, 1
 if not rs.eof then
	strImage = rs("theImage").GetChunk(1024000)
    Response.ContentType = "image/jpg"
    Response.BinaryWrite strImage
	
	end if
rs.close
conn.close
%>
