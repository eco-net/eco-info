<!--#include virtual="/Connections/ecoinfo.asp" -->
<%Dim sql,sqlreplace
sqlreplace="SELECT m.emailaddress,m.phone,"
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
<title>print liste</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td> 
<div align="center"> 
<p class="contentHeader1"><img src="/shared/showimg.asp?id=78"><br>
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
<td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
<td rowspan="6"><%if rsPageData.Fields.Item("phone").Value<>"" then response.write rsPageData.Fields.Item("phone").Value end if %></td>
</tr>
<% 
While NOT rsPageData.EOF 

response.write "<tr><td><img src=""/shared/graphics/spacer.gif"" width=""3"" height=""4""></td></tr>"
response.write "<tr><td><span class=""contentheader2"">"
response.write rsPageData.Fields.Item("title").Value
response.write "</span>"
if rsPageData.Fields.Item("subtitle").Value<>"" then
response.write "<br><span class=""listheaderItalic"">" & rsPageData.Fields.Item("subtitle").Value & "</span>"
end if 
if rsPageData.Fields.Item("shortdescription").Value<>"" then
response.write "<br>" & replace(rsPageData.Fields.Item("shortdescription").Value & "",chr(13),"<br>")
end if 
response.write "<span class=""listheader""><br>"
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
IF rsPageData.Fields.Item("city").value<>"" THEN response.write " | " & rsPageData.Fields.Item("city").value
response.write "</span>"
response.write "</td></tr>"
if rsPageData.Fields.Item("organizers").value<>"" THEN 
response.write "<tr><td>Arrangør: " & rsPageData.Fields.Item("organizers").value & "</td></tr>"
end if
response.write "<tr><td><img src=""/shared/graphics/spacer.gif"" width=""3"" height=""7""></td></tr>"
response.write "<tr><td background=""/shared/graphics/layout/dots_horz.gif"" height=""1""><img src=""/shared/graphics/spacer.gif"" width=""3"" height=""1""></td></tr>"


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