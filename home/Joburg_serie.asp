<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="99%" class="sidebarHeader">&nbsp;&nbsp;JOBURG BILLEDER</td>
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
<td>
<p><img src="/shared/graphics/spacer.gif" width="3" height="8"></p>
<p class="listheader">Klik p&aring; billederne for at se en stor st&oslash;rrelse</p>
</td>
</tr>
</table>
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
<p><br>
<span class="contentHeader1"> 16 stemningsbillder fra WSSD i Johannesburg</span></p>
<p> <img src="/home/serie/thumbnails/IMG_0451.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0451.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Ungdomskor fra Soveto (South West Township) paa Nasrec (Peoples Global Forum)</p>
<p><img src="/home/serie/thumbnails/IMG_0469.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0469.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Kendt koreansk folkemusiker giver gribende fredskoncert paa Nasrec (Peoples Global 
Forum)</p>
<p><img src="/home/serie/thumbnails/IMG_0472.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0472.jpg','Joburg','width=620,height=480')" hspace="25"><br>
En stor gruppe koreanere hylde deres musik idol da han gav fredskoncert paa Nasrec 
(Peoples Global Forum)</p>
<p><img src="/home/serie/thumbnails/IMG_0502.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0502.jpg','Joburg','width=620,height=480')" hspace="25"> 
<br>
Greening the summit</p>
<p><img src="/home/serie/thumbnails/IMG_0474.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0474.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Greening the summit </p>
<p><img src="/home/serie/thumbnails/IMG_0499.jpg" width="56" height="75" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0499.jpg','Joburg','width=620,height=820')" hspace="25"><br>
USA er ikke det mest populare land - kollagen indeholder et element fra 92-gruppens 
Frihedsgudinde plakat!</p>
<p><br>
<img src="/home/serie/thumbnails/IMG_0505.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0505.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Sandton hvor det officielle topmode holdes</p>
<p><img src="/home/serie/thumbnails/IMG_0519.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0519.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Miljominister Hans Chr. Schmidt giver briefing for den danske delegation paa Hotel 
Balalajka</p>
<p><br>
<img src="/home/serie/thumbnails/IMG_0534.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0534.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Alexandre township - derfra hvor Global Peoples March gik lordag d. 31. august\</p>
<p><br>
<img src="/home/serie/thumbnails/IMG_0539.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0539.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Sydafrikas president holder tale foer Global Peoples March lordag d. 31. august 
paa Alexandre sportsstadion.</p>
<p><img src="/home/serie/thumbnails/IMG_0547.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0547.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Global Peoples March, lordag d. 31. august</p>
<p><br>
<img src="/home/serie/thumbnails/IMG_0569.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0569.jpg','Joburg','width=620,height=480')" hspace="25"><br>
To danske hojskole folk og Oko-net udveklser synspunkter med statsminister Anders 
Fogh Rasmussen ved dennes reception for danske delegerede sondag d. 1. September 
paa et Hotel I Sandton</p>
<p><br>
<img src="/home/serie/thumbnails/IMG_0586.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0586.jpg','Joburg','width=620,height=480')" hspace="25"> 
<br>
Kofi Annan, FN's generalsekretaer holdt tale Mandag d. 2. September paa Golbal 
Peoples Forum</p>
<p><br>
<img src="/home/serie/thumbnails/IMG_0587.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0587.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Happening fra Friends of the Earth - Whos voice is missing!</p>
<p><br>
<img src="/home/serie/thumbnails/IMG_501.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0501.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Fra Ubunto hvor alle nationer har en stand - her er det den danske stand hvor 
man bl.a. kan finde haefter med den nationale danske strategi for baeredygtig 
udvikling.</p>
<p><img src="/home/serie/thumbnails/IMG_554.jpg" width="75" height="56" onClick="MM_openBrWindow('/home/image.asp?img=IMG_0554.jpg','Joburg','width=620,height=480')" hspace="25"><br>
Den danske 92-gruppe holder strategi mode paa Hotel Balalajka I Sandton - I midten 
er det koordinator John Nordbo fra 92-gruppen.<br>
</p>
<p><br>
<br>
</p>
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
