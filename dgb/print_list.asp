<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%Dim sql,sqlreplace
sqlreplace="SELECT m.isbn,m.editor,m.editoremail,m.editorwww,"
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
<title>&Oslash;ko-info - Det Gr&oslash;nne Bibliotek om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td width="60%"> 
<div align="center"> 
<p class="contentHeader1"> 
<!--#include virtual="/shared/showimage.asp" -->
<img src="<%=getImage(25,"R")%>"> 
<br>
www.eco-info.dk </p>
<p>&nbsp;</p>
</div>
</td>
<td valign="top"> 
<div align="center"> 
<p><span class="contentHeader1">Det Gr&oslash;nne Bibliotek<br>
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
<table border="0" cellpadding="0" cellspacing="0" width="90%" align="center">
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="2"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
</tr>
<tr> 
<td colspan="2">&nbsp;</td>
</tr>
<tr> 
<td colspan="2">&nbsp;</td>
</tr>
<% 
While NOT rsPageData.EOF 
%>
<tr> 
<td width="60%" valign="top"> <span class="contentHeader2"><%=rsPageData.Fields.Item("title").Value%></span><br>
<%if rsPageData.Fields.Item("subtitle").Value<>"" then%>
<span class="subtitle"><%=rsPageData.Fields.Item("subtitle").Value%></span><span class="listheaderItalic"><br>
</span> 
<%end if %>
<%if rsPageData.Fields.Item("shortdescription").Value<>"" then%>
<%=rsPageData.Fields.Item("shortdescription").Value%><br>
<%end if %>
<%if rsPageData.Fields.Item("author").value<>"" THEN %>
<span class="listheader">Forfatter:</span>&nbsp; <%=rsPageData.Fields.Item("author").Value%><br>
<%end if %>
</td>
<td valign="top"> 
<% if rsPageData.Fields.Item("editor").value<>"" THEN %>
<span class="listheader"> Udgiver/Forlag:</span> <%=rsPageData.Fields.Item("editor").value %><br>
<%end if %>
<%if rsPageData.Fields.Item("editoremail").Value<>"" then %>
<span class="listheader">Forlagets Email:</span> <a href="mailto:<%=rsPageData.Fields.Item("editoremail").Value %>"><%=rsPageData.Fields.Item("editoremail").Value %></a><br>
<%end if %>
<%if rsPageData.Fields.Item("isbn").Value<>"" then %>
<span class="listheader">ISBN:</span> <%=rsPageData.Fields.Item("isbn").Value %><br>
<%end if %>
<%if rsPageData.Fields.Item("webaddress").Value<>"" then %>
<span class="listheader">www:</span> <a href="http://<%=rsPageData.Fields.Item("webaddress").Value %>" target="_blank"><%=rsPageData.Fields.Item("webaddress").Value %></a><br>
<%end if %>
</td>
</tr>
<tr> 
<td colspan="2">&nbsp;</td>
</tr>
<tr> 
<td colspan="2">&nbsp;</td>
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
