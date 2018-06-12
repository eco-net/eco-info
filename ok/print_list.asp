<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
SetLocale("da")
Dim sql,sqlreplace
sqlreplace="SELECT m.emailaddress,m.phone,m.fulladdress,m.postnr,"
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
<title>&Oslash;ko-info - &Oslash;ko-Kalenderen om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
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
<p><span class="contentHeader1">&Oslash;ko-kalenderen<br>
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
<%if rsPageData.Fields.Item("organizers").value<>"" THEN %>
Arrangør:&nbsp; <%=rsPageData.Fields.Item("organizers").Value%><br>
<%end if %>
</td>
<td valign="top"> <span class="listheader"> 
<%
IF DatePart("m",rsPageData.Fields.Item("startdate").value)=DatePart("m",rsPageData.Fields.Item("enddate").value) THEN
	IF DatePart("d",rsPageData.Fields.Item("startdate").value)=DatePart("d",rsPageData.Fields.Item("enddate").value) THEN
		response.write FormatDateTime(rsPageData.Fields.Item("startdate").value,vbLongDate)
	ELSE
		response.write DatePart("d",rsPageData.Fields.Item("startdate").value) & " til " & FormatDateTime(rsPageData.Fields.Item("enddate").value,vbLongDate)
	END IF
ELSEIF rsPageData.Fields.Item("enddate").value<>"" THEN
	response.write FormatDateTime(rsPageData.Fields.Item("startdate").value,vbLongDate) & " til " & FormatDateTime(rsPageData.Fields.Item("enddate").value,vbLongDate)
ELSE
	response.write FormatDateTime(rsPageData.Fields.Item("startdate").value,vbLongDate)
END IF

%>
</span> <br>
<% if rsPageData.Fields.Item("fulladdress").value<>"" THEN %>
<%=rsPageData.Fields.Item("fulladdress").value %><br>
<%end if %>
<%
IF rsPageData.Fields.Item("city").value<>"" THEN %>
<%= rsPageData.Fields.Item("postnr").value %>&nbsp;<%=rsPageData.Fields.Item("city").value%><br>
<% end if %>
<%if rsPageData.Fields.Item("phone").Value<>"" then %>
<span class="listheader">Telefon:</span> <%=rsPageData.Fields.Item("phone").Value %><br>
<%end if %>
<%if rsPageData.Fields.Item("emailaddress").Value<>"" then %>
<span class="listheader">Email:</span> <a href="mailto:<%=rsPageData.Fields.Item("emailaddress").Value %>"><%=rsPageData.Fields.Item("emailaddress").Value %></a><br>
<%end if %>
<%if rsPageData.Fields.Item("www").Value<>"" then %>
<span class="listheader">www:</span> <a href="http://<%=rsPageData.Fields.Item("www").Value %>" target="_blank"><%=rsPageData.Fields.Item("www").Value %></a><br>
<%end if %>
</td>
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
