<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include file="act_sqlcalc.asp" -->
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT m.subtitle,m.fulladdress,m.id,m.title,m.shortdescription,m.www,m.startdate,m.enddate,m.organizers,p.city  FROM (eical_maindata m LEFT JOIN eisys_postnr p ON m.postnr=p.postnr)  WHERE 0=0  ORDER BY m.startdate"
rsPageData.Source=replace(rsPageData.Source,"0=0",theWhere)
rsPageData.Source=replace(rsPageData.Source,"(eical_maindata m LEFT JOIN eisys_postnr p ON m.postnr=p.postnr)",theFrom)
rsPageData.Source=replace(rsPageData.Source,"ORDER BY m.startdate",theOrder)

rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
'response.write rsPageData.Fields.Item("id").value
'response.end
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 10
Dim Repeat1__index
Repeat1__index = 0
rsPageData_numRows = rsPageData_numRows + Repeat1__numRows
%>
<!--#include virtual="/shared/inc_listnavigation.asp"-->
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>&Oslash;ko-info - &Oslash;ko-Kalenderen om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
<!-- #EndEditable --> 
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="7" marginwidth="0" marginheight="7">
<table width="752" border="0" cellspacing="0" cellpadding="0" align="center" name="Pagetable">
<tr> 
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
<td width="750" valign="top"> 
<!-- START HEADER --><!-- #BeginLibraryItem "/Library/header.lbi" -->
<table width="750" border="0" cellspacing="0" cellpadding="0" name="Header">
<form name="headerform" method="post" action="/home/searchall.asp" onSubmit="doSearch(document.headerform);">
<tr> 
<td valign="top" rowspan="3"><img src="/shared/graphics/logo.gif" alt="&Oslash;ko-info er Danmarks &oslash;kologiske portal, med alt om &oslash;kologi, milj&oslash; og b&aelig;redygtig udvikling." width="180" height="33"> 
<table width="180" border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td width="8"><br>
</td>
<td class="sitetag"> Guide til det gr&oslash;nne Danmark - vejen til &oslash;kologisk information!<br>
</td>
<td width="8"><br>
</td>
</tr>
</table></td>
<td valign="top" width="1"><br>
</td>
<td height="17" align="right" colspan="3"> <a href="/home/index.asp">Forside</a> 
| <a href="/home/about_1.asp">Om �ko-info</a> | <a href="/home/searching.asp">S&oslash;getips</a> 
| <a href="/home/advertising.asp">For annonc&oslash;rer</a> | <a href="/home/insider.htm">Kom 
med i De Gr&oslash;nne Sider</a> </td>
</tr>
<tr> 
<td valign="top" rowspan="2" width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="3"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td width="388" height="64"> 
<div align="center">
<!--#include virtual="/log/ei/banner.asp"-->
</div>
</td>
<td background="/shared/graphics/layout/dots_vert.gif" width="1"><br>
</td>
<td width="180" align="center" valign="top" background="/shared/graphics/layout/search_bkgrd.gif"> 
<table width="150" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
<tr> 
<td><span class="formLabeltext">S&oslash;g i</span> 
<select name="thesection" class="formInputobject">
<option value="2">De Gr&oslash;nne Sider</option>
<option value="3">&Oslash;ko-kalenderen</option>
<option value="4">Det Gr&oslash;nne Bibliotek</option>
<option value="5">Opslagstavlen</option>
</select>
<span class="formLabeltext">... efter ordet</span> 
<input type="text" name="thetext" class="formInputobject" value="<skriv tekst og tast enter>" onFocus="this.value=''" onChange="doSearch(document.headerform);document.headerform.submit();" size="20">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
<!-- #EndLibraryItem --><!-- END HEADER -->
<!-- #BeginEditable "menu" --> 
<%DIM curtab
curtab=2%>
<!-- #BeginLibraryItem "/Library/menu.lbi" --><table width="750" border="0" cellspacing="0" cellpadding="0" name="Menu">
<tr>
<%
'-- tab1 --
response.write "<td><img src=""/shared/graphics/menu/stretch.gif"" width=""65"" height=""22"">"
IF curtab=1 THEN
	response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
		"<a href=""/dgs/index.asp""><img src=""/shared/graphics/menu/dgs_on.gif"" width=""92"" height=""22"" border=""0""></a>"
ELSE
	response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">" &_
		"<a href=""/dgs/index.asp""><img src=""/shared/graphics/menu/dgs_off.gif"" width=""92"" height=""22"" border=""0""></a>"
END IF
'-- tab2 --
IF curtab=2 THEN
	response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
		"<a href=""/ok/index.asp""><img src=""/shared/graphics/menu/ok_on.gif"" width=""94"" height=""22"" border=""0""></a>"
ELSE
	IF curtab=1 THEN
		response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
	ELSE
		response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
	END IF
	response.write "<a href=""/ok/index.asp""><img src=""/shared/graphics/menu/ok_off.gif"" width=""94"" height=""22"" border=""0""></a>"
END IF
'-- tab3 --
IF curtab=3 THEN
	response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
		"<a href=""/dgb/index.asp""><img src=""/shared/graphics/menu/dgb_on.gif"" width=""121"" height=""22"" border=""0""></a>"
ELSE
	IF curtab=2 THEN
		response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
	ELSE
		response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
	END IF
	response.write "<a href=""/dgb/index.asp""><img src=""/shared/graphics/menu/dgb_off.gif"" width=""121"" height=""22"" border=""0""></a>"
END IF
'-- tab4 --
IF curtab=4 THEN
	response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
		"<a href=""/opsl/index.asp""><img src=""/shared/graphics/menu/opsl_on.gif"" width=""86"" height=""22"" border=""0""></a>"
ELSE
	IF curtab=3 THEN
		response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
	ELSE
		response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
	END IF
	response.write "<a href=""/opsl/index.asp""><img src=""/shared/graphics/menu/opsl_off.gif"" width=""86"" height=""22"" border=""0""></a>"
END IF
'-- tab5 --
IF curtab=5 THEN
	response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
		"<a href=""/news/index.asp""><img src=""/shared/graphics/menu/news_on.gif"" width=""52"" height=""22"" border=""0""></a>" &_
		"<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
ELSE
	IF curtab=4 THEN
		response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
	ELSE
		response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
	END IF
	response.write "<a href=""/news/index.asp""><img src=""/shared/graphics/menu/news_off.gif"" width=""52"" height=""22"" border=""0""></a>" &_
		"<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
END IF
'-- tab6 --
'IF curtab=6 THEN
'	response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
'		"<img src=""/shared/graphics/menu/tema_on.gif"" width=""30"" height=""22"">" &_
'		"<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
'ELSE
'	IF curtab=5 THEN
'		response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
'	ELSE
'	END IF
'	response.write "<a href=""/tema/index.asp""><img src=""/shared/graphics/menu/tema_off.gif"" width=""30"" height=""22"" border=""0""></a>" &_
'		"<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
'END IF
response.write "<img src=""/shared/graphics/menu/stretch.gif"" width=""66"" height=""22""></td>"
%>
</tr>
<%IF curtab>0 THEN%>
<tr><td class="menuBar">
<%
SET fs = CreateObject("Scripting.FileSystemObject")
Set ts=fs.OpenTextFile(Server.mapPath("inc_submenu.txt"))
response.write ts.ReadAll()
ts=""
fs=""
%>
</td></tr>
<%END IF%>
<tr><td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr></table><!-- #EndLibraryItem --><!-- #EndEditable --> 
<table width="750" border="0" cellspacing="0" cellpadding="0" name="Contentarea">
<tr> 
<td width="180" valign="top" name="leftbar" bgcolor="#ECECDD"><!-- #BeginEditable "leftbar" --> 
<!--#include file="inc_search.asp"-->
<%dim editpage,editid
editpage=""
editid=""
%>
<!--#include virtual="/admin/inc_mypages.asp"-->

<table width="90%" border="0" cellspacing="0" cellpadding="10" align="center">
<form name="form2" method="post" action="print_list.asp" target="_blank">
<tr> 
<td> 
<div align="center">
<p> 
<input type="hidden" name="searchdesc" value="<%=searchDescr%>">
<input type="hidden" name="sql" value="<%=rsPageData.Source%>">
<br>
Din s&oslash;gning kan udskrives med alle v&aelig;sentlige data:<br>
<input type="submit" name="Submit2" value="Se Udskrift" class="formSubmitbutton">
</p>
</div>
</td>
</tr>
</form>
</table>
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" --> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top" width="98%"> 
<% If Not rsPageData.EOF Or Not rsPageData.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="8" height="30"></td>
<td class="contentHeader1" colspan="2" width="98%">Fandt <%=rsPageData_total%> arrangementer</td>
<td><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
</tr>
<tr> 
<td colspan="4" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr valign="top"> 
<td><img src="/shared/graphics/spacer.gif" width="8" height="30"></td>
<td><%=recInfo%></td>
<td align="right" nowrap><%=naviHtml%></td>
<td><br>
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td colspan="3"><img src="/shared/graphics/spacer.gif" width="1" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td valign="top" width="99%"> <br>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsPageData.EOF)) 

response.write "<tr><td><img src=""/shared/graphics/spacer.gif"" width=""3"" height=""4""></td></tr>"
response.write "<tr><td><span class=""contentheader2""><a href=""" & replace(detaillink,"##",CStr(MM_offset+Repeat1__index)) & """  title=""Viser alle oplysninger om dette arrangement"">"
response.write rsPageData.Fields.Item("title").Value
response.write "</a></span>"
if rsPageData.Fields.Item("subtitle").Value<>"" then
response.write "<br><span class=""subtitle"">" & rsPageData.Fields.Item("subtitle").Value & "</span>"
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
response.write "<tr><td>Arrang�r: " & rsPageData.Fields.Item("organizers").value & "</td></tr>"
end if
response.write "<tr><td><img src=""/shared/graphics/spacer.gif"" width=""3"" height=""7""></td></tr>"
response.write "<tr><td background=""/shared/graphics/layout/dots_horz.gif"" height=""1""><img src=""/shared/graphics/spacer.gif"" width=""3"" height=""1""></td></tr>"

  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsPageData.MoveNext()
Wend
%>
</table>
<br>
<br>
</td>
<td><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
</tr>
<tr> 
<td><br>
</td>
<td valign="top" align="right"><%=naviHtml%></td>
<td><br>
</td>
</tr>
</table>
<% End If ' end Not rsPageData.EOF Or NOT rsPageData.BOF %>
<% If rsPageData.EOF And rsPageData.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="8"><br>
</td>
<td class="contentHeader1" height="22">Fandt ingen arrangementer</td>
<td width="8"><br>
</td>
</tr>
<tr> 
<td colspan="3" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><br>
</td>
<td> <br>
<br>
<br>
Der er ingen arrangementer i &Oslash;ko-kalenderen, der matcher din s�gning.<br>
Det skyldes sansynligvis, at du har lavet en meget pr&aelig;cis s&oslash;gning. 
</td>
<td><br>
</td>
</tr>
</table>
<% End If ' end rsPageData.EOF And rsPageData.BOF %>
</td>
<td valign="top" height="100%" bgcolor="#ECECDD"> 
<!--#include virtual="/log/ei/ok/inc_annoncer.asp"-->
</td>
</tr>
</table>
<!-- #EndEditable --> </td>
</tr>
<tr> 
<td valign="bottom" align="left"> 
<!--#include virtual="/shared/pagetools.asp"-->
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
</tr>
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="3"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr align="center"> 
<td colspan="3" class="footer" height="20" valign="middle">
<!--#include virtual="/shared/footer.asp"-->
</td>
</tr>
</table>
</body>
<!-- #EndTemplate --></html>
<%
rsPageData.Close()
%>
<!--include virtual="/shared/log.asp"-->
<%'reg("oklist")%>