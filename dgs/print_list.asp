<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%Dim sql,sqlreplace
sqlreplace="SELECT m.emailaddress1,m.phone1,m.phone2,m.street1,m.zip_dk,m.city,"
sql=request("sql")
sql=replace(sql,"SELECT",sqlreplace)
'response.write(sql)
'response.end
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source=sql
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0

%>
<html>
<head>
<title>&Oslash;ko-info - De Gr&oslash;nne Sider om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td width="60%"> 
<div align="center"> 
<p class="contentHeader1"> 
<!--#include file="../shared/showimage.asp"-->
<img src="<%=getImage(25,"R")%>"> 
<br>
www.eco-info.dk </p>
<p>&nbsp;</p>
</div>
</td>
<td valign="top"> 
<div align="center"> 
<p><span class="contentHeader1">De Gr&oslash;nne Sider<br>
</span><%=date%><span class="contentHeader1"><br>
<br>
</span>s&oslash;gning: 
<%response.write(request("searchdesc"))%>
<br>
</p>
</div>
</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="5" width="100%" align="center">
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="5"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
</tr>
<tr> 
<td colspan="5">&nbsp;</td>
</tr>
<tr> 
<td class="listheader" valign="top" width="20%"> <span class="listheader">Organisation 
</span></td>
<td valign="top" class="listheader" width="20%">Kontaktperson</td>
<td valign="top" class="listheader" width="20%">Adresse</td>
<td valign="top" class="listheader" width="15%">Telefon</td>
<td valign="top" class="listheader" width="25%">Net</td>
</tr>
<tr> 
<td colspan="5">&nbsp;</td>
</tr>
<tr> 
<td colspan="5">&nbsp;</td>
</tr>
<% 
While NOT rsPageData.EOF 
%>
<tr> 
<td class="listheader" valign="top"><%=rsPageData("name")%></td>
<td valign="top"><%=rsPageData("firstname")%>&nbsp;<%=rsPageData("middlename")%>&nbsp;<%=rsPageData("lastname")%></td>
<td valign="top"><%=rsPageData("street1")%><br>
<%=rsPageData("zip_dk")%>&nbsp;<%=rsPageData("city")%></td>
<td valign="top"><%=rsPageData("phone1")%>
<%if rsPageData("phone2")<>"" then %>
<br>
<%=rsPageData("phone2")%>
<%end if %>
</td>
<td valign="top"><a href="mailto:<%=rsPageData("emailaddress1")%>"><%=rsPageData("emailaddress1")%></a><br>
<a href="http://<%=rsPageData("www")%>" target="_blank"><%=rsPageData("www")%></a></td>
</tr>
<tr> 
<td colspan="5">&nbsp;</td>
</tr>
<tr> 
<td colspan="5">&nbsp;</td>
</tr>
<%
 	rsPageData.MoveNext()
Wend
	rsPageData.Close()
%>
</table>
</body>
</html>
<script language="javascript">
print();
</script>
