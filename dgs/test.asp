<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<% response.write (request("new") & "<br>") 
DIM theFrom, theWhere, theOrder, searchDescr, eicount, eibackcount,einextcount
theFrom="eiorg_maindata m"
theWhere="(m.validated=1 or m.validated=0)  AND m.stopped=0"

'theWhere="validated=1"
theOrder="ORDER BY m.name"
searchDescr=""

IF LEN(request("new"))>0 THEN
	theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date),DatePart("d",date))
	theDate=DateAdd("d",(0-request("new")),theDate)
	theWhere=theWhere & " AND m.createdate>=" & FormatDateDB(theDate)
	searchDescr="oprettet indenfor de seneste <b>" & request("new") & " dage</b>"
	response.write (theDate & "<br>")
	response.write (theWhere & "<br>")
	response.write (searchDescr & "<br>")
End if
FUNCTION FormatDateDB(theDate)
	theDate=CDate(theDate)
	FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
		right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
END FUNCTION
%>
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT m.id,m.zip,m.name,m.shortdescription,m.www,m.firstname,m.middlename,m.lastname,m.fullname,m.stopped  FROM eiorg_maindata m  WHERE 0=0  ORDER BY m.name"
rsPageData.Source=replace(rsPageData.Source,"0=0",theWhere)
rsPageData.Source=replace(rsPageData.Source,"eiorg_maindata m",theFrom)
rsPageData.Source=replace(rsPageData.Source,"ORDER BY m.name",theOrder)
'response.write rsPageData.Source
'response.end
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<!--#include virtual="/shared/inc_detailnavigation.asp" -->

<%

while not rsPageData.EOF
response.write "<tr><td><span class=""contentheader2""><a href=""" & replace(detaillink,"##",CStr(MM_offset+Repeat1__index)) & """  title=""" & replace(rsPageData.Fields.Item("shortdescription").Value,chr(13),"<br>") & """>"
response.write rsPageData("id") & " - "
response.write rsPageData("name") & " - "
response.write rsPageData("fullname") & " <br> "

rsPageData.MoveNext
wend
%>
</body>
</html>
