<%@LANGUAGE="VBSCRIPT"%>
<!--include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include file="act_sqlcalc.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT m.*  FROM eiorg_maindata m  WHERE 0=0 ORDER BY m.name"
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
<!--#include file="act_detail.asp" -->
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>&Oslash;ko-info - De Gr&oslash;nne Sider om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
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
curtab=1%>
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
<!--#include virtual="/dgs/inc_search.asp"-->
<%dim editpage,editid
editpage="s"
editid=rsPageData("id")
%>
<!--#include virtual="/admin/inc_mypages.asp"-->
<p>&nbsp;</p>
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" --> 
<table width="100%" border="0" cellspacing="0" cellpadding="16" background="/dgs/graphics/detail_header_dgs_bkgrd.gif" name="detailHeader">
<tr> 
<td valign="top"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif" name="detailContents">
<tr> 
<td class="contentHeader1" align="left"> 
<%IF LEN(rsPageData.Fields.Item("name").Value)=0 THEN
	response.write rsPageData.Fields.Item("firstname").Value & "&nbsp;" & rsPageData.Fields.Item("middlename").Value & "&nbsp;" & rsPageData.Fields.Item("lastname").Value
ELSE
	response.write rsPageData.Fields.Item("name").Value
END IF%>
<%IF LEN(request.cookies("eiuserid"))>0 THEN%>
&nbsp;&nbsp;&nbsp;<a href="http://insider.eco-info.dk/dgs/edit.asp?id=<%=(rsPageData.Fields.Item("id").Value)%>" class="sitetag" target="_blank">Rediger 
kortet</a> 
<%END IF%>
</td>
<td>&nbsp;</td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td height"20"><%=recInfo%></td>
<td height="20" align="right" nowrap><%=naviHTML%></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="1" height="5"></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> 
<table border="0" cellspacing="0" cellpadding="0" width="100%">
<tr> 
<td class="contentHeader2">Beskrivelse</td>
</tr>
<tr> 
<td valign="top"> 
<p><%=replace(rsPageData.Fields.Item("description").Value & "",chr(13),"<br>")%></p>
<span class="contentHeader2">Andre oplysninger</span><br>
<%=catList & sbjList & okList %>
<%=dgbList & courseList%> </td>
</tr>
</table>
</td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td valign="top"> 
<table border="0" cellspacing="0" cellpadding="0">
<tr> 
<td class="contentHeader2">Fakta</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><span class="faktaboksheader"><br>
Adresse<br>
</span> 
<%IF LEN(rsPageData.Fields.Item("name").Value)=0 THEN
	response.write rsPageData.Fields.Item("firstname").Value & "&nbsp;" & rsPageData.Fields.Item("lastname").Value
ELSE
	response.write rsPageData.Fields.Item("name").Value
END IF%>
<br>
<% if rsPageData.Fields.Item("adrco").Value <>"" then %>
<%if left(rsPageData("adrco"),3)="c/o" or left(rsPageData("adrco"),4)=" c/o" then response.write "" else response.write "c/o:" %><%=rsPageData.Fields.Item("adrco").Value%><br>
<%END IF%>
<%if rsPageData.Fields.Item("street1").Value <>"" then%>
<%=(rsPageData.Fields.Item("street1").Value)%>
<%end if%>
<%if rsPageData.Fields.Item("street2").Value<>"" then%>
<br> <nobr><%=(rsPageData.Fields.Item("street2").Value)%>
</nobr> 
<%end if %>
<%if Len(rsPageData.Fields.Item("zip").Value)>1 then %>
<br><nobr><%=(rsPageData.Fields.Item("zip").Value)%> &nbsp;<%=(rsPageData.Fields.Item("city").Value)%></nobr><br>
<% end if %>
<% if rsPageData.Fields.Item("country").Value <> "Danmark" then %>
<%=(rsPageData.Fields.Item("country").Value)%> <br>
<%end if %>
<%if rsPageData.Fields.Item("firstname").Value <>"" or rsPageData.Fields.Item("middlename").Value <>"" or rsPageData.Fields.Item("lastname").Value <>"" then %>
<br>
<span class="listheader">Kontaktperson</span><br>
<%if rsPageData.Fields.Item("title").Value <>"" then%>
<%=(rsPageData.Fields.Item("title").Value)%><br>
<%end if%>
<%=(rsPageData.Fields.Item("firstname").Value)%>&nbsp;<%=(rsPageData.Fields.Item("middlename").Value)%>&nbsp;<%=(rsPageData.Fields.Item("lastname").Value)%> <br>
<%end if%>
<%if rsPageData.Fields.Item("phone1").Value <>"" then%>
<br>
<span class="faktaboksheader">Telefon</span> <br>
<%=splitPhone((rsPageData.Fields.Item("phone1").Value))%> 
<%end if%>
<%if rsPageData.Fields.Item("phone2").Value <>"" then%>
<br>
<%=splitPhone((rsPageData.Fields.Item("phone2").Value))%> 
<%end if%>
<%if rsPageData.Fields.Item("phone3").Value <>"" then%>
<br><%=splitPhone((rsPageData.Fields.Item("phone3").Value))%> 
<%end if%>
<%if rsPageData.Fields.Item("fax").Value <>"" then%>
<br>
<span class="faktaboksheader"><br>
Fax</span><br>
<%=splitPhone((rsPageData.Fields.Item("fax").Value))%> 
<%end if%>
<% if rsPageData.Fields.Item("emailaddress1").Value<>"" or  rsPageData.Fields.Item("emailaddress2").Value<>"" then%>
<span class="faktaboksheader"><br>
<br>
E-mail</span><br>
<a href="mailto:<%=(rsPageData.Fields.Item("emailaddress1").Value)%>" title="�bner en ny mail i din mail-klient med denne adresse som modtager"><%=(rsPageData.Fields.Item("emailaddress1").Value)%></a> 
<% if rsPageData.Fields.Item("emailaddress2").Value<>"" then %>
<br>
<a href="mailto:<%=(rsPageData.Fields.Item("emailaddress2").Value)%>" title="�bner en ny mail i din mail-klient med denne adresse som modtager"><%=(rsPageData.Fields.Item("emailaddress2").Value)%></a> 
<%end if %>
<%end if %>
<%if rsPageData.Fields.Item("www").Value <>"" then%>
<br>
<span class="faktaboksheader"><br>
Hjemmeside</span><br>
<a href="http://<%=(rsPageData.Fields.Item("www").Value)%>" title="�bner hjemmesiden i et nyt vindue" target="_blank"><%=rsPageData.Fields.Item("www").Value%><img src="/shared/graphics/layout/arrows_fwd.gif" width="10" height="7" border="0"></a>
<% if rsPageData.Fields.Item("www2").Value<>"" then%>
<br><a href="http://<%=(rsPageData.Fields.Item("www2").Value)%>" title="�bner hjemmesiden i et nyt vindue" target="_blank"><%=rsPageData.Fields.Item("www2").Value%><img src="/shared/graphics/layout/arrows_fwd.gif" width="10" height="7" border="0"></a> 
<%end if%>
<%end if%>
</td>
</tr>
</table>
</td>
</tr>
</table>
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
