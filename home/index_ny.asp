<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="shared/common.asp" -->
<%
set rsNews = Server.CreateObject("ADODB.Recordset")
rsNews.ActiveConnection = MM_ecoinfo_STRING
rsNews.Source = "SELECT top 5 Artikel_ID, CreateDate, Header  FROM bu_Artikel  ORDER BY CreateDate desc"
rsNews.CursorType = 0
rsNews.CursorLocation = 2
rsNews.LockType = 3
rsNews.Open()
rsNews_numRows = 0
%>
<%
set rsAfstemning = Server.CreateObject("ADODB.Recordset")
rsAfstemning.ActiveConnection = MM_ecoinfo_STRING
rsAfstemning.Source = "SELECT top 10 n.id, n.forslag, n.stemmer,n.descr,e.name  FROM eiafstem_ning n LEFT JOIN eiafstem_emne e ON n.emne_id=e.id  WHERE e.chosen=1  ORDER BY n.stemmer desc"
rsAfstemning.CursorType = 0
rsAfstemning.CursorLocation = 2
rsAfstemning.LockType = 3
rsAfstemning.Open()
rsAfstemning_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsNews_numRows = rsNews_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
rsAfstemning_numRows = rsAfstemning_numRows + Repeat2__numRows
%>
<html>
<head>
<title>&Oslash;ko-info</title>
<meta name="keywords" content="�kologi, milj�, b�redygtig udvikling, natur, gr�n, landbrug, f�devarer, agenda 21, energi, by�kologi, forurening, �kologisk, ecology, b�redygtighed, agenda21, bes�gslandbrug, g�rdbutik, gr�n guide, naturvejleder, genbrug, klima, �kosamfund, trafik, �kologisk byggeri, sundhed, helse, helsekost, permakultur, alternativ, vand, atomkraft">
<meta name="description" content="�ko-info er Danmarks �kologiske portal, med alt om �kologi, milj� og b�redygtig udvikling.">
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
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
<!--#include virtual="/shared/stat.asp" -->
<%
DIM curtab
curtab=0%><!-- #BeginLibraryItem "/Library/menu.lbi" --><table width="750" border="0" cellspacing="0" cellpadding="0" name="Menu">
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
</tr></table><!-- #EndLibraryItem --><table width="750" border="0" cellspacing="0" cellpadding="0" name="Contentarea">
<tr> 
<td width="180" valign="top" name="leftbar" bgcolor="#ECECDD"> 
<p> 
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader"> &nbsp;&nbsp;AKTUELT</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr> 
<td width="90%"> 
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsNews.EOF)) 
%>
<a href="/news/detail.asp?id=<%=(rsNews.Fields.Item("Artikel_ID").Value)%>"><span class="listheader"><%=toString(rsNews.Fields.Item("Header").Value)%></span></a><br>
<%theDate=(rsNews.Fields.Item("CreateDate").Value)
dato=datepart("d",theDate) & "-" & datepart("m",theDate) & "-" & datepart("yyyy",theDate)%>
<%=dato%> 
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsNews.MoveNext()
	if not rsNews.eof then response.write "<br><br>"
Wend
%>
</td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader"> &nbsp;&nbsp;KOM MED... </td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr> 
<td class="sidebarHeader"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td> 
<p>Gr&oslash;nne organisationer - NGO'ere som offentlige - og private personer 
kan oprette og pr&aelig;sentere sig selv i databasen 
<!-- - med mulighed for at tilknytte 
arrangementer til &Oslash;ko-kalenderen og publikationer i Det Gr&oslash;nne Bibliotek -->
<br>
<br>
<a href="#">G&aring; til &Oslash;ko-Info Insider</a></p>
</td>
</tr>
</table>
</td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader"> &nbsp;&nbsp;ANDRE &Oslash;KO-NET SITES</td>
</tr>
<tr> 
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<td colspan="2"> 
<div align="center"> 
<p><a href="#"><br>
B�redygtig Udvikling</a><br>
<br>
<a href="#">&Oslash;ko-net </a></p>
<p>&nbsp;</p>
</div>
</td>
</tr>
</table>
<p align="center"><br>
</p>
</td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="388" height="100%" valign="top" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> 
<!--#include virtual="log/ei/home/inc_leader.txt"-->
<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader"> &nbsp;&nbsp;SEKTIONER</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
<!--#include virtual="log/ei/home/inc_sections.txt"-->
</td>
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
<td width="180" valign="top" name="rightbar" bgcolor="#ECECDD"> 
<!--#include file="rightpane.asp"-->
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;NYHEDSBREV</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td colspan="2" class="sidebarText"> 
<script src="/shared/newsmail.js"></script>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<form name="newsmail" action="" onSubmit="return newsMail(1,document.newsmail.themail.value,document.newsmail.site.value);">
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr> 
<td><span class="sidebarHeader">Din emailadresse </span> 
<input type="text" name="themail" value="" class="formInputobject">
<input type="hidden" value="1" name="site">
<input type="button" value="Afmeld" class="formbutton"  name="afmeld" onClick="javascript:newsMail(2,document.newsmail.themail.value,document.newsmail.site.value);">
<input type="submit" value="Tilmeld" name="tilmeld" class="formSubmitbutton">
</td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</form>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;AFSTEMNING</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td class="listheader" colspan="2"><%=(rsAfstemning.Fields.Item("name").Value)%></td>
</tr>
<tr> 
<td><i>emne </i></td>
<td><i>stemmer </i></td>
</tr>
<% If Not rsAfstemning.EOF Or Not rsAfstemning.BOF Then %>
<% 
While ((Repeat2__numRows <> 0) AND (NOT rsAfstemning.EOF)) 
%>
<tr> 
<td width="75%"> 
<%=(rsAfstemning.Fields.Item("forslag").Value)%> </td>
<td> 
<div align="right"><%=(rsAfstemning.Fields.Item("stemmer").Value)%></div>
</td>
</tr>
<% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsAfstemning.MoveNext()
Wend
%>
<% End If ' end Not rsAfstemning.EOF Or NOT rsAfstemning.BOF %>
</table>
<p align="center"> 
<input type="button" name="b1" value="Deltag" class="formSubmitbutton" onClick="MM_openBrWindow('stem.asp','Afstemning','scrollbars=yes,resizable=yes,width=400,height=600')">
<br>
<br>
</p>
</td>
</tr>
</table>
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
<td colspan="3" class="footer" height="20" valign="middle">&Oslash;ko-info er 
etableret af <a href="http://www.eco-net.dk" target="_blank">&Oslash;ko-net</a>, 
der st&oslash;ttes af <a href="http://www.mst.dk/gronfond/" target="_blank">Den 
Gr&oslash;nne Fond</a>.<br>
En anden portal fra &Oslash;ko-net er <a href="http://www.baeredygtigudvikling.dk" target="_blank">www.B&aelig;redygtigUdvikling.nu</a><br>
<br>
</td>
</tr>
</table>
</body>
</html>
<%
rsNews.Close()
%>
<%
rsAfstemning.Close()
%>
<!--#include virtual="/shared/log.asp"-->

