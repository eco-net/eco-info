<!--#include virtual="/shared/showimage.asp" -->
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>&Oslash;ko-info - guide til alt om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
<!-- #EndEditable --> 
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script></head>
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
| <a href="/home/about_1.asp">Om Øko-info</a> | <a href="/home/searching.asp">S&oslash;getips</a> 
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
<!-- #BeginEditable "menu" --> <!-- #BeginLibraryItem "/Library/menu.lbi" --><table width="750" border="0" cellspacing="0" cellpadding="0" name="Menu">
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
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td> 
<p>&nbsp;</p>
<p>&nbsp;</p>
<p class="sidebarHeader">Bestil folder og plakat <br>
p&aring; tlf. 62244324</p>
<p class="sidebarHeader"><a href="graphics/Folder_EcoInfo.pdf" target="_blank">Hent 
Folder som pdf</a> (1,22MB)</p>
<p class="sidebarHeader"><a href="graphics/Folder_EcoInfo_light.pdf" target="_blank">Hent 
Folder som pdf</a> (308KB)</p>
<p class="sidebarHeader"><a href="graphics/Plakat_EcoInfo.pdf" target="_blank">Hent 
Plakat som pdf</a> (251KB)</p>
<p>For at downloade pdf'erne kan du h&oslash;jreklikke p&aring; linket og v&aelig;lge: 
Gem destination som...</p>
<p>For at l&aelig;se og printe pdf'ere skal du bruge Acrobat Reader.</p>
<p><a href="http://www.adobe.com/products/acrobat/readstep2.html">Hent Acrobat 
Reader Here</a></p>
<p class="sidebarHeader">&nbsp;</p>
</td>
</tr>
</table>
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
<td colspan="2" class="contentHeader1" align="left">&Oslash;ko-info kampagnemateriale</td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td height"20">K&aelig;re &Oslash;ko-info bruger.</td>
<td height="20" align="right" nowrap>13-11-2002</td>
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
<td class="contentHeader2">&nbsp;</td>
</tr>
<tr> 
<td valign="top"> 
<p><br>
&Oslash;ko-net har fremstillet foldere og plakater som kampagnemateriale for denne 
site, &Oslash;ko-info.</p>
<p>Foldere og plakater er delt ud til samtlige skoler i Danmark og til mange af 
vores samarbejdspartnere, gr&oslash;nne og folkelige organisationer, for at udbrede 
kendskabet til dette samlingssted for milj&oslash; og &oslash;kologi-interesserede 
i Danmark.</p>
<p>Det er vores ydmyge h&aring;b at tilfredse brugere, der &oslash;nsker at &Oslash;ko-info 
skal best&aring;, trods de for vort land s&aring; trange tider, vil hj&aelig;lpe 
med at udbrede kendskabet til &Oslash;ko-info.</p>
<p>Du kan bestille folder og plakat og hj&aelig;lpe med distributionen. Ring 62 
24 24 43 eller mail: <a href="mailto:eco-net@eco-net.dk">eco-net@eco-net.dk</a><br>
Gode steder at anbringe materialet kan v&aelig;re:<br>
Biblioteker<br>
Uddannelsessteder<br>
Supermarkeder<br>
Helsekostbutikker <br>
IT-brugersteder<br>
Medborgerhuse<br>
Universiteternes instituter<br>
<br>
</p>
<p class="listheader">Tak til f&oslash;lgende, der har bidraget til folderen med 
et foto:</p>
<table width="100%" border="0" cellspacing="5" cellpadding="0">
<tr valign="top"> 
<td>Fotograf<br>
<span class="listheader">Ulla Tr&aelig;dmark Jensen</span><br>
H&aelig;stvej 23<br>
8380 Trige</td>
<td>Tlf. priv.: 86 98 98 95<br>
Mobil tlf.: 28 40 73 30 <br>
e-mail: <a href="mailto:ullatj@post.tele.dk">ullatj@post.tele.dk</a></td>
</tr>
<tr valign="top"> 
<td><span class="listheader">ph7 kommunikation</span><br>
v. Peder Hovgaard<br>
Studsgade 7, 2<br>
8000 &Aring;rhus C</td>
<td>Tlf.: 86 76 04 16<br>
Fax: 86 76 57 67<br>
e-mail: <a href="mailto:ph@ph7.dk">ph@ph7.dk</a><br>
Hjemmeside: <a href="http://www.ph7.dk" target="_blank">www.ph7.dk</a></td>
</tr>
<tr valign="top"> 
<td><span class="listheader">Mads Izmodenov Eskesen</span><br>
K&oslash;benhavns Milj&oslash;- og Energikontor - KMEK<br>
Blegdamsvej 4 B, st<br>
2200 K&oslash;benhavn N</td>
<td>Tlf. arb.: 35 37 36 36<br>
Fax: 35 37 36 76<br>
e-mail: <a href="mailto:mads@ove.org">mads@ove.org</a></td>
</tr>
<tr valign="top"> 
<td><span class="listheader">Allan Born&oslash; Clausen &amp; Ida Dahl Nielsen</span><br>
Hegnstrup &oslash;kologiske jordbrug<br>
Gl. K&oslash;benhavnsvej 14<br>
3550 Slangerup</td>
<td>Tlf.: 47 33 44 20 <br>
e-mail: <a href="mailto:hegnstrup@bigfoot.com">hegnstrup@bigfoot.com</a></td>
</tr>
<tr valign="top"> 
<td><span class="listheader">Landsforeningen Levende Hav</span><br>
Juelsg&aring;rdvej 27, Ferring<br>
7620 Lemvig</td>
<td>Tlf.: 97 89 54 55<br>
Mobil-tlf.: 51 24 57 12<br>
Fax: 97 89 56 55<br>
e-mail: <a href="mailto:llh@levende-hav.dk">llh@levende-hav.dk</a><br>
Hjemmeside: <a href="http://www.levende-hav.dk" target="_blank">www.levende-hav.dk</a></td>
</tr>
<tr valign="top"> 
<td><span class="listheader">Portalen Naturnet.dk</span><br>
Skov- og Naturstyrelsen, Friluftskontoret<br>
Haraldsgade 53<br>
2100 K&oslash;benhavn &Oslash;</td>
<td>Tlf.: 39 47 23 08<br>
Fax: 39 47 21 65<br>
e-mail: <a href="mailto:naturnet@sns.dk">naturnet@sns.dk</a><br>
Hjemmeside: <a href="http://www.naturnet.dk" target="_blank">www.naturnet.dk</a></td>
</tr>
</table>
</td>
</tr>
</table>
</td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td valign="top" width="200"> 
<table border="0" cellspacing="0" cellpadding="0">
<tr> 
<td class="contentHeader2">&nbsp;</td>
</tr>
<tr> 
<td valign="top"> 
<div align="center"></div>
</td>
</tr>
</table>
<div align="center">klik og se stor st&oslash;rrelse:<br>
<br>
</div>
<div align="center"> 
<p><img src="<%=getImage(282,"R")%>" onClick="MM_openBrWindow('kampagne_img.asp?img=1','kampagne','resizable=yes,width=870,height=620')"></p>
<p><img src="<%=getImage(283,"R")%>" onClick="MM_openBrWindow('kampagne_img.asp?img=2','kampagne','resizable=yes,width=870,height=620')"></p>
<p><img src="<%=getImage(285,"R")%>" onClick="MM_openBrWindow('kampagne_img.asp?img=3','kampagne','resizable=yes,width=550,height=770')"></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p class="listheader">&nbsp;</p>
</div>
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
<!--#include virtual="/shared/log.asp"-->
