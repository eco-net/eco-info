<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/showimage.asp" -->
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM bestyrelse_maindata  ORDER BY orden"
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsPageData_numRows = rsPageData_numRows + Repeat1__numRows
%>
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>&Oslash;ko-info - guide til alt om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
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
<%level1=6%>
<!--#include file="inc_about_leftbar.asp"-->
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" -->
  <table width="100%" border="0" cellspacing="0" cellpadding="16" background="/dgs/graphics/detail_header_dgs_bkgrd.gif" name="detailHeader">
    <tr>
      <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif" name="detailContents">
          <tr>
            <td colspan="2" class="contentHeader1" align="left">Lokalafdelinger i &Oslash;ko-net:</td>
          </tr>
          <tr>
            <td colspan="2" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
          </tr>
          <tr>
            <td height"20">&nbsp;</td>
            <td height="20" align="right" nowrap>&nbsp;</td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="/shared/graphics/spacer.gif" width="1" height="5"></td>
            </tr>
          </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td valign="top"><table border="0" cellspacing="0" cellpadding="0" width="100%">
                  <tr>
                    <td><p>&Oslash;ko-net varmer op til  flere landsd&aelig;kkende aktiviteter og har fra 2007 oprette/stiftet lokalafdelinger  af &Oslash;ko-net i de nye regioner i Danmark.</p>
                        <p>De fire nye lokal afdelinger er:</p>
                      <p><a href="hovedstaden.asp">&Oslash;ko-net Hovedstaden</a><br>
                            <a href="sjaelland.asp">&Oslash;ko-net Sj&aelig;lland</a><br>
                            <a href="midtjylland.asp">&Oslash;ko-net Midtjylland</a><br>
                            <a href="nordjylland.asp">&Oslash;ko-net Nordjylland</a><br>
                            <br>
                        Og formodentlig f&oslash;lger i for&aring;ret &Oslash;ko-net  Syddanmark.</p>
                      <p>Med udgangspunkt i den nye struktur udvikling i  Danmark &ndash; den nye kommunalreformen og de nye regioner i Danmark fra d. 1.  januar 2007 - har &Oslash;ko-net nu taget initiativ til at blive forankret mere fysisk  i det nye Danmark. Dette er sket ved at &Oslash;ko-net har taget initiativ til at der  oprettes lokale afdelinger af &Oslash;ko-net, der alle har f&oslash;lgende form&aring;l:</p>
                      <p>&rdquo;Form&aring;let  med &Oslash;ko-net XXXXX er p&aring; et folkeligt niveau at informere, oplyse og inspirere  omkring natur og milj&oslash;, &oslash;kologi og b&aelig;redygtig udvikling - samt at skabe debat  og netv&aelig;rk omkring &oslash;kologiske og b&aelig;redygtige tiltag.<br>
                        Ved &oslash;kologi  forst&aring;s en husholdning med ressourcer, der er i balance med naturen.<br>
                        Ved b&aelig;redygtig  udvikling forst&aring;s en udvikling, der tager globale, milj&oslash;m&aelig;ssige og sociale  hensyn b&aring;de til nulevende og kommende generationer.<br>
                        Der l&aelig;gges  v&aelig;gt p&aring; ogs&aring; at engagere ungdommen i foreningen og i arbejdet med ovenn&aelig;vnte  form&aring;l&rdquo;.</p>
                      <p>I december 2006 er der i f&oslash;rste omgang stiftet  de fire f&oslash;rste lokalafdelinger af &Oslash;ko-net &ndash; hhv. &Oslash;ko-net Nordjylland, &Oslash;ko-net  Midtjylland, &Oslash;ko-net Sj&aelig;lland og &Oslash;ko-net Hovedstaden. Vi arbejder ogs&aring; p&aring; at  stifte et &Oslash;ko-net Syddanmark, men i f&oslash;rste omgang varetages aktiviteterne her af  &Oslash;ko-nets landsafdelingen i Ollerup ved Svendborg.</p>
                      <p>I 2007 er m&aring;let at der minimum afholdes et  arrangement i hver region &ndash; og &Oslash;ko-net vil samtidig arbejde p&aring; at skabe mere  opm&aelig;rksomhed omkring &Oslash;ko-nets aktiviteter og m&aelig;rkesager. Et andet m&aring;l er at  blive flere medlemmer &ndash; gerne et godt stykke over 1000 medlemmer i 2007 &ndash; fra  600 i 2006.</p>
                      <p>Som medlem af &Oslash;ko-net st&oslash;tter man et  landsd&aelig;kkende arbejde for:</p>
                      <ul>
                        <ul>
                          <li>at virkeligg&oslash;re Uddannelse  for B&aelig;redygtig Udvikling (UBU) i Danmark og Balanceakten som et visuelt  international symbol for FN-ti&aring;ret for UBU</li>
                          <li>at etablere en  landsd&aelig;kkende klimanetv&aelig;rk for unge til unge og unge til &aelig;ldre</li>
                          <li>at skabe et landsd&aelig;kkende  netv&aelig;rk af lokale &Oslash;ko-net aktiviteter</li>
                          <li>at udvikle Danmarks gr&oslash;nne  svar p&aring; de gule sider &ndash; De Gr&oslash;nne Sider p&aring; nettet</li>
                          <li>at afholde et &aring;rligt  &Oslash;ko-net landsseminar med folkeoplysning, debat og meninger, mad og musik &amp;  &oslash;kologi og kultur</li>
                          <li>m.m, m.m, m.m...</li>
                        </ul></td>
                  </tr>
                  <tr>
                    <td valign="top"><p>&nbsp;</p>
                        <p><br>
                      </p></td>
                  </tr>
              </table></td>
              <td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
              <td width="1" background="/shared/graphics/layout/dots_vert.gif"><br></td>
              <td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
              <td valign="top" width="200"><table border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="188" class="contentHeader2"><strong>Kontakt information </strong>: </td>
                  </tr>
                  <tr>
                    <td><a href="hovedstaden.asp">Hovedstaden</a><br>
                        <a href="sjaelland.asp">Sj&aelig;lland</a><br>
                        <a href="midtjylland.asp">Midtjylland</a><br>
                        <a href="nordjylland.asp">Nordjylland</a><br>
                        <br></td>
                  </tr>
                  <tr>
                    <td><strong>Vedt&aelig;gter:</strong></td>
                  </tr>
                  <tr>
                    <td><a href="hovedstaden_vedt.asp">Hovedstaden</a><br>
                        <a href="sjaelland_vedt.asp">Sj&aelig;lland</a><br>
                        <a href="midtjylland_vedt.asp">Midtjylland</a><br>
                        <a href="nordjylland_vedt.asp">Nordjylland</a><br>
                        <br>
                        <img src="/shared/graphics/danmark_nye_regioner_ny.gif" width="480" height="384" border="0" usemap="#Map">
                        <map name="Map">
                          <area shape="poly" coords="79,103,132,103,164,79,139,52,99,70,83,81" href="Nordjylland.asp">
                          <area shape="poly" coords="43,169,102,170,121,183,113,211,89,217,73,199,45,209,39,192" href="Midtjylland.asp">
                          <area shape="poly" coords="44,246,72,240,88,253,115,257,117,269,122,283,116,303,54,289,39,263" href="#">
                          <area shape="poly" coords="225,265,236,246,247,247,253,258,263,265,269,277,272,304,257,317,232,309" href="Sjaelland.asp">
                          <area shape="poly" coords="262,210,280,236,354,246,354,225,272,201" href="Hovedstaden.asp">
                        </map></td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td valign="top">&nbsp;</td>
              <td>&nbsp;</td>
              <td background="/shared/graphics/layout/dots_vert.gif">&nbsp;</td>
              <td>&nbsp;</td>
              <td valign="top">&nbsp;</td>
            </tr>
        </table></td>
    </tr>
  </table>
  <br>
  <br>
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
