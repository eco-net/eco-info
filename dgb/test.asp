<html><!-- #BeginTemplate "/Templates/3cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
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
curtab=3%>
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
<!--#include file="inc_about.txt"-->
<!--#include file="inc_search.asp"-->
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="388" height="100%" valign="top" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" -->
<table width="388" border="0" cellspacing="0" cellpadding="16" background="/dgb/graphics/sub_index_header_dgb_bkgrd.gif" name="subIndexHeader">
<tr> 
<td valign="top"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
<tr> 
<td colspan="2" class="contentHeader1">Emneord</td>
</tr>
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="2"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td width="66%" valign="top">
<p>
Nedenfor ser du en oversigt over emneordene p&aring; &Oslash;ko-info.<br>
Klik p&aring; en hoved- eller underkategori for at f&aring; vist  publikationer med det p&aring;g&aelig;ldende emneord.
</p>
</td>
<td width="34%">&nbsp;</td>
</tr>
</table>
</td>
</tr>
</table>
<table width="90%" border="0" cellspacing="0" align="center">
<tr> 
<td valign="top">
<table width='100%' border='0' cellspacing='0' cellpadding='0' align='center'>
<tr>
<td colspan='3'><img src='/shared/graphics/spacer.gif' width='3' height='8'></td>
</tr>
<tr>
<td colspan='3' background='/shared/graphics/layout/dots_horz.gif'><img src='/shared/graphics/spacer.gif' width='1' height='1' border='0'></td>
</tr>
<tr>
<td colspan='3'><img src='/shared/graphics/spacer.gif' width='1' height='4' border='0'></td>
</tr>
<tr>
<td colspan='3'><span class='contentHeader2'><a href='list.asp?dgbsbj=x0.5'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>Andet</a></span><br>
Emneord der ikke h�rer hjemme i en hovedgruppe.</td>
</tr>
<tr>
<td colspan='3'><img src='/shared/graphics/spacer.gif' width='1' height='4' border='0'></td>
</tr>
<tr valign='top'>
<td width='49%'><a href='list.asp?dgbsbj=308&dgbsbjname=Affald / Genbrug'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>Affald / Genbrug</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
</td>
<td width='8'><img src='/shared/graphics/spacer.gif' width='8' height='1'></td>
<td width='49%'><a href='list.asp?dgbsbj=321&dgbsbjname=Arbejde'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>Arbejde</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
<a href='list.asp?dgbsbj=310&dgbsbjname=B�redygtig udvikling'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>B�redygtig udvikling</a><br>
</td>
</tr>
<tr>
<td colspan='3'><img src='/shared/graphics/spacer.gif' width='1' height='4' border='0'></td>
</tr>
<tr>
<td colspan='3'><img src='/shared/graphics/spacer.gif' width='1' height='4' border='0'></td>
</tr>
<tr>
<td colspan='3' background='/shared/graphics/layout/dots_horz.gif'><img src='/shared/graphics/spacer.gif' width='1' height='1' border='0'></td>
</tr>
<tr>
<td colspan='3'><img src='/shared/graphics/spacer.gif' width='1' height='4' border='0'></td>
</tr>
<tr>
<td colspan='3'><span class='contentHeader2'><a href='list.asp?dgbsbj=x1&dgbsbjname=emneordsgruppe1'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>emneordsgruppe1</a></span><br>
emneordsgruppe1</td>
</tr>
<tr>
<td colspan='3'><img src='/shared/graphics/spacer.gif' width='1' height='4' border='0'></td>
</tr>
<td width='49%'><a href='list.asp?dgbsbj=115&dgbsbjname=Agenda 21'><img src='/shared/graphics/layout/arrows_lead.gif' width='10' height='7' border='0'>Agenda 21</a></td>
</tr>
<tr>
<td colspan='3'><img src='/shared/graphics/spacer.gif' width='1' height='4' border='0'></td>
</tr>
<tr>
<td colspan='3'><img src='/shared/graphics/spacer.gif' width='3' height='8'></td>
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
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="180" valign="top" name="rightbar" bgcolor="#ECECDD"><!-- #BeginEditable "rightbar" -->
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr>
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr>
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;VED DU NOGET ?</td>
</tr>
<tr>
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr>
<td>Har du kendskab til en god artitkel, bog, hjemmeside eller andet s&aring; 
<a href="#">fort&aelig;l os om det.</a><br>
<br>
Se ogs&aring; <a href="http://insider.eco-info.dk">&Oslash;ko-info Insider</a> om dine muligheder for at bruge Det Gr&oslash;nne Biblioteket p� din egen site.</td>
</tr>
<tr>
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
<!--#include virtual="/log/ei/dgb/inc_annoncer.asp"-->
<!-- #EndEditable --></td>
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
<!--#include virtual="/shared/log.asp"-->