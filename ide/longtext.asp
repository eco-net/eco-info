<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="file://///Server/www/Connections/ecoinfo.asp" -->
<%
Dim rsPageData__theID
rsPageData__theID = "0"
if (request("id") <> "") then rsPageData__theID = request("id")
%>
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT descr, name, email, title  FROM buidekuvoese_maindata  WHERE id=" + Replace(rsPageData__theID, "'", "''") + ""
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td class="listheader"><%=(rsPageData.Fields.Item("title").Value)%></td>
<td> 
<div align="right"><%=(rsPageData.Fields.Item("name").Value)%></div>
</td>
</tr>
<tr> 
<td colspan="2"><%=(rsPageData.Fields.Item("descr").Value)%></td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<input type="button" name="Button" value="Luk Vindue" class="formInputobject" onClick="window.close();">
</div>
</td>
</tr>
</table>
</body>

</html>
<%
rsPageData.Close()
%>
