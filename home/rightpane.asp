
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
FUNCTION FormatDateDB(theDate)
	theDate=CDate(theDate)
	FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
		right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
END FUNCTION
%>
<%
SetLocale("da")
DIM uge, maaned, rs, weekorg, monthorg, weekok, monthok, weekdgb, monthdgb, weekopsl, monthopsl, weeknews, monthnews 
	theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date),DatePart("d",date))
	uge=DateAdd("d",-7,theDate)
	uge=FormatDateDB(uge)
	maaned=DateAdd("d",-30,theDate)
	maaned=FormatDateDB(maaned)
	
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_ecoinfo_STRING
	rs.CursorType = 0
	rs.CursorLocation = 2
	rs.LockType = 3
	'organisationer
	rs.Source = "SELECT Count(*) as theCount  FROM eiorg_maindata WHERE createdate>=" & uge 
	rs.Open()
	weekorg = rs.Fields.Item("theCount").value
	rs.Close()
	rs.Source = "SELECT Count(*) as theCount  FROM eiorg_maindata WHERE createdate>=" & maaned
	rs.Open()
	monthorg = rs.Fields.Item("theCount").value
	rs.Close()
	'kalender
	rs.Source = "SELECT Count(*) as theCount  FROM eical_maindata WHERE createdate>=" & uge 
	rs.Open()
	weekok= rs.Fields.Item("theCount").value
	rs.Close()
	rs.Source = "SELECT Count(*) as theCount  FROM eical_maindata WHERE createdate>=" & maaned
	rs.Open()
	monthok = rs.Fields.Item("theCount").value
	rs.Close()
	'bibliotek
	rs.Source = "SELECT Count(*) as theCount  FROM eilib_maindata WHERE createdate>=" & uge 
	rs.Open()
	weekdgb= rs.Fields.Item("theCount").value
	rs.Close()
	rs.Source = "SELECT Count(*) as theCount  FROM eilib_maindata WHERE createdate>=" & maaned
	rs.Open()
	monthdgb = rs.Fields.Item("theCount").value
	rs.Close()
	'opslagstavle
	rs.Source = "SELECT Count(*) as theCount  FROM eiopsl_maindata WHERE createdate>=" & uge 
	rs.Open()
	weekopsl= rs.Fields.Item("theCount").value
	rs.Close()
	rs.Source = "SELECT Count(*) as theCount  FROM eiopsl_maindata WHERE createdate>=" & maaned
	rs.Open()
	monthopsl = rs.Fields.Item("theCount").value
	rs.Close()
	'nyheder
	rs.Source = "SELECT Count(*) as theCount  FROM bu_Artikel WHERE create_date>=" & uge 
	rs.Open()
	weeknews= rs.Fields.Item("theCount").value
	rs.Close()
	rs.Source = "SELECT Count(*) as theCount  FROM bu_Artikel WHERE create_date>=" & maaned
	rs.Open()
	monthnews = rs.Fields.Item("theCount").value
	rs.Close()
	
%>
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;HOLD DIG AJOUR</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr valign="top"> 
<td colspan="2" class="sidebarText"> 
<table width="90%" border="0" cellspacing="0" cellpadding="2" align="center">
<tr> 
<td colspan="2"><img src="/shared/graphics/spacer.gif" width="3" height="3"></td>
</tr>
<tr> 
<td colspan="2" class="sidebarHeader">Oprettet den seneste uge</td>
</tr>
<tr> 
<td> <a href="/dgs/list.asp?new=7" title="Se organisationer oprettet de seneste 7 dage">De 
Gr&oslash;nne Sider</a> </td>
<td nowrap>
<div align="right"><%=weekorg%></div>
</tr>
<tr> 
<td> <a href="/ok/list.asp?new=7" title="Se arrangementer oprettet de seneste 7 dage">&Oslash;ko-Kalenderen</a> 
</td>
<td nowrap>
<div align="right"><%=weekok%></div>
</td>
</tr>
<tr> 
<td> <a href="/dgb/list.asp?new=7" title="Se publikationer oprettet de seneste 7 dage">Det 
Gr&oslash;nne Bibliotek</a></td>
<td nowrap>
<div align="right"><%=weekdgb%></div>
</td>
</tr>
<tr> 
<td> <a href="/opsl/list.asp?new=7" title="Se opslag oprettet de seneste 7 dage">Opslagstavlen</a> 
</td>
<td nowrap>
<div align="right"><%=weekopsl%></div>
</td>
</tr>
<tr> 
<td> <a href="/news/list.asp?new=7" title="Se nyheder oprettet de seneste 7 dage">Aktuelt</a> 
</td>
<td nowrap>
<div align="right"><%=weeknews%></div>
</td>
</tr>
</table>
<table width="90%" border="0" cellspacing="0" cellpadding="2" align="center">
<tr> 
<td colspan="2" class="sidebarHeader">Oprettet den seneste m&aring;ned</td>
</tr>
<tr> 
<td> <a href="/dgs/list.asp?sektion=dgs&new=30" title="Se organisationer oprettet de seneste 30 dage">De 
Gr&oslash;nne Sider</a> </td>
<td nowrap>
<div align="right"><%=monthorg%></div>
</td>
</tr>
<tr> 
<td> <a href="/ok/list.asp?new=30" title="Se arrangementer oprettet de seneste 30 dage">&Oslash;ko-Kalenderen</a> 
</td>
<td nowrap>
<div align="right"><%=monthok%></div>
</td>
</tr>
<tr> 
<td> <a href="/dgb/list.asp?new=30" title="Se publikationer oprettet de seneste 30 dage">Det 
Gr&oslash;nne Bibliotek</a></td>
<td nowrap>
<div align="right"><%=monthdgb%></div>
</td>
</tr>
<tr> 
<td> <a href="/opsl/list.asp?new=30" title="Se organisationer oprettet de seneste 30 dage">Opslagstavlen</a> 
</td>
<td nowrap>
<div align="right"><%=monthopsl%></div>
</td>
</tr>
<tr> 
<td> <a href="/news/list.asp?new=30" title="Se nyheder oprettet de seneste 30 dage">Aktuelt</a> 
</td>
<td nowrap>
<div align="right"><%=monthnews%></div>
</td>
</tr>
<tr> 
<td colspan="2"><img src="/shared/graphics/spacer.gif" width="3" height="3"></td>
</tr>
</table>
</td>
</tr>
</table>

