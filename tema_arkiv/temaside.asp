<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsTema__theID
rsTema__theID = "0"
if (request("id") <> "") then rsTema__theID = request("id")
theWhere="id=" + Replace(rsTema__theID, "'", "''") + ""
%>
<%
set rsTema = Server.CreateObject("ADODB.Recordset")
rsTema.ActiveConnection = MM_ecoinfo_STRING
rsTema.Source = "SELECT id,header,subheader  FROM eitema_maindata  WHERE  0=0"
if request("id")<>"" then
rsTema.Source = replace(rsTema.Source,"WHERE 0=0",theWhere)
end if
rsTema.CursorType = 0
rsTema.CursorLocation = 2
rsTema.LockType = 3
rsTema.Open()
rsTema_numRows = 0
%>
<%
Dim rsArtikel__theID
rsArtikel__theID = rsTema("id")
'if (rsTema("id") <> "") then rsArtikel__theID = rsTema("id")
%>
<%
set rsArtikel = Server.CreateObject("ADODB.Recordset")
rsArtikel.ActiveConnection = MM_ecoinfo_STRING
rsArtikel.Source = "SELECT *  FROM eitema_artikel  WHERE tema_id=" + Replace(rsArtikel__theID, "'", "''") + "  ORDER BY place"
if request("artikel_id")<>"" then
rsArtikel.Source = "SELECT *  FROM eitema_artikel  WHERE id=" & request("artikel_id")
end if
rsArtikel.CursorType = 0
rsArtikel.CursorLocation = 2
rsArtikel.LockType = 3
rsArtikel.Open()
rsArtikel_numRows = 0
%>
<%
Dim rsArtikler__theID
rsArtikler__theID = "0"
if (rsTema("id")<> "") then rsArtikler__theID = rsTema("id") %>
<%
set rsArtikler = Server.CreateObject("ADODB.Recordset")
rsArtikler.ActiveConnection = MM_ecoinfo_STRING
rsArtikler.Source = "SELECT id, title  FROM eitema_artikel  WHERE tema_id=" + Replace(rsArtikler__theID, "'", "''") + "  ORDER BY place"
rsArtikler.CursorType = 0
rsArtikler.CursorLocation = 2
rsArtikler.LockType = 3
rsArtikler.Open()
rsArtikler_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsArtikler_numRows = rsArtikler_numRows + Repeat1__numRows
%>
<!--#include virtual="shared/common.asp" -->
<html>
<head>
<title>&Oslash;ko-info - Tema-arkiv om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
<meta name="keywords" content="�kologi, milj�, b�redygtig udvikling, natur, gr�n, landbrug, f�devarer, agenda 21, energi, by�kologi, forurening, �kologisk, ecology, b�redygtighed, agenda21, bes�gslandbrug, g�rdbutik, gr�n guide, naturvejleder, genbrug, klima, �kosamfund, trafik, �kologisk byggeri, sundhed, helse, helsekost, permakultur, alternativ, vand, atomkraft">
<meta name="description" content="�ko-info er Danmarks �kologiske portal, med alt om �kologi, milj� og b�redygtig udvikling.">
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
<%
DIM curtab
curtab=0%>
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
</tr></table><!-- #EndLibraryItem --><table width="750" border="0" cellspacing="0" cellpadding="0" name="Contentarea">
<tr> 
<td width="180" valign="top" name="leftbar" bgcolor="#ECECDD"> 
<p> 
<p align="center"><br>
</p>
</td>
<td width="388" height="100%" valign="top" name="maincontent"> 
<table width="90%" height="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td valign="top"> 
<div align="center"><span class="contentHeader1"><br>
<%=(rsTema.Fields.Item("header").Value)%></span><br>
<%=(rsTema.Fields.Item("subheader").Value)%> </div>
</td>
</tr>
<tr> 
<td valign="top"> 
<p><span class="contentHeader2"><br>
Tema Artikler<br>
</span> 
<% If Not rsArtikler.EOF Or Not rsArtikler.BOF Then %>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsArtikler.EOF)) 
%>
<a href="temaside.asp?artikel_id=<%=(rsArtikler.Fields.Item("id").Value)%>"><span class="listheader"><%=(rsArtikler.Fields.Item("title").Value)%></span> </a><br>
<br>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsArtikler.MoveNext()
Wend
%>
<% End If ' end Not rsArtikler.EOF Or NOT rsArtikler.BOF %>
</p>
<p class="contentHeader2">&nbsp;</p>
</td>
</tr>
<tr> 
<td valign="top"> 
<%rsArtikel.MoveFirst %>
<p class="contentHeader2"><%=(rsArtikel.Fields.Item("title").Value)%></p>
<p><%=(rsArtikel.Fields.Item("tekst").Value)%> 
<% if rsArtikel.Fields.Item("link").Value<>"" then %>
<br>
<br>
<a href="<%=rsArtikel.Fields.Item("link").Value%>" target=_blank> 
<% =rsArtikel.Fields.Item("linktext").Value %>
</a> 
<% end if %>
</p>
<p>&nbsp;</p>
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
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td colspan="2" class="sidebarText">&nbsp; </td>
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
rsTema.Close()
%>
<%
rsArtikel.Close()
%>
<%
rsArtikler.Close()
%>
<!--#include virtual="/shared/log.asp"-->

