<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="shared/common.asp" -->
<!--#include virtual="log/ei/home/newsnr.inc"-->
<%
function picasaImgSize(imgsrc,w)
i=1
imgname=Right(imgsrc,i) 
while InStr(imgname,"/")=0 
i=i+1
imgname=Right(imgsrc,i) 
wend
imgsrc=Left(imgsrc,Len(imgsrc)-i)
if Right(imgsrc,5)="/s144" then 'hvis thumbnail er kopieret
imgsrc=Left(imgsrc,Len(imgsrc)-5)
end if
imgsrc=imgsrc & w & imgname
picasaImgSize=imgsrc
end function

set rsNews = Server.CreateObject("ADODB.Recordset")
rsNews.ActiveConnection = MM_ecoinfo_STRING
rsNews.Source = "SELECT top " & number & " a.Artikel_ID, a.create_date, a.Header ,a.img_id,a.imagefilename1,a.img_text FROM bu_Artikel a LEFT JOIN bu_Artikel_site s ON a.Artikel_ID=s.artikel_id  WHERE s.econetsite<>1 and  a.header<>'' ORDER BY a.create_date desc"
rsNews.CursorType = 0
rsNews.CursorLocation = 2
rsNews.LockType = 3
rsNews.Open()
rsNews_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsNews_numRows = rsNews_numRows + Repeat1__numRows
%>
<html>
<head>
<title>&Oslash;ko-info - guide til alt om &oslash;kologi, natur, milj&oslash; 
og b&aelig;redygtighed</title>
<meta name="keywords" content="økologi, miljø, bæredygtig udvikling, natur, grøn, landbrug, fødevarer, agenda 21, energi, byøkologi, forurening, økologisk, ecology, bæredygtighed, agenda21, besøgslandbrug, gårdbutik, grøn guide, naturvejleder, genbrug, klima, økosamfund, trafik, økologisk byggeri, sundhed, helse, helsekost, permakultur, alternativ, vand, atomkraft">
<meta name="description" content="Øko-info er Danmarks økologiske portal, med alt om økologi, miljø og bæredygtig udvikling.">
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="7" marginwidth="0" marginheight="7" onLoad="MM_preloadImages('/shared/graphics/pagetools/print_txt.gif','/shared/graphics/pagetools/bookmark_txt.gif','/shared/graphics/pagetools/mail_txt.gif')">
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
<!--include virtual="/shared/stat.asp" -->
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
</tr></table><!-- #EndLibraryItem --><!--#include virtual="/shared/showimage.asp" -->
<table width="750" border="0" cellspacing="0" cellpadding="0" name="Contentarea">
<tr> 
<td width="180" valign="top" name="leftbar" bgcolor="#ECECDD"> 
<!-- sitemap -->
<!-- /Sitemap -->
<!-- Nyeste i Aktuelt -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td></tr>
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22" border="0"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;NYESTE I AKTUELT</td>
</tr>
<tr><td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td></tr>
<tr><td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td></tr>
<tr><td> 
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsNews.EOF)) 
%><center>
<a href="/news/detail.asp?id=<%=(rsNews.Fields.Item("Artikel_ID").Value)%>"><span class="listheader"><%=toString(rsNews.Fields.Item("Header").Value)%></span></a><br>
<%theDate=(rsNews.Fields.Item("create_date").Value)
dato=datepart("d",theDate) & "-" & datepart("m",theDate) & "-" & datepart("yyyy",theDate)%>
<%=dato%> 
<br>
<% if rsNews("img_id")>0 then %>
<a href="/news/detail.asp?id=<%=(rsNews.Fields.Item("Artikel_ID").Value)%>" title="<%=dato%>">
<img src="<%=getImage(rsNews("img_id"),"R") %>" width="75" border="0"></a>
 <% end if %>
 <% if Len(rsNews.Fields.Item("imagefilename1").Value)>0 then %>
<a href="/news/detail.asp?id=<%=rsNews("Artikel_ID")%>">
<img src="<%=picasaImgSize(rsNews("imagefilename1"),"/s144")%>" alt="<%=rsNews("img_text")%>" width=60 hspace="0" vspace="0" border=0 /></a> 
<% end if %>
</center>
     
       
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsNews.MoveNext()
	if not rsNews.eof then response.write "<br><br>"
Wend
%>
      
      <br>
        <br>  <div align="center">
        <a href="/news/index.asp">Tidligere</a></div>
</td></tr>
<tr><td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td></tr>
</table>
</td></tr>
</table>
<!-- /Nyeste i aktuelt -->

<!-- Støt øko-info -->
<!-- / Støt Øko-info -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;KOM MED...</td>
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

<%dim editpage,editid
editpage=""
editid=""
%>
<!--#include virtual="/admin/inc_mypages.asp"-->

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
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;ANDRE &Oslash;KO-NET SITES</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<td colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr> 
<td> <a href="http://www.baeredygtigudvikling.nu">Bæredygtig Udvikling</a> </td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr> 
<td> <a href="http://www.sustainabledevelopment.dk">Sustainable Development</a> 
</td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr> 
<td> <a href="http://www.eco-net.dk">&Oslash;ko-net</a> </td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
  </tr>
  <tr bgcolor="#FAFAF4">
    <td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
    <td width="98%" class="sidebarHeader">&nbsp;&nbsp;S&Oslash;G P&Aring;  &Oslash;KO-NET SITES</td>
  </tr>
  <tr>
    <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
  </tr>
  <td colspan="2"><p align="center">
            <script src="http://gmodules.com/ig/ifr?url=http://www.google.com/coop/api/015918078326015659760/cse/vy5mj_xc020/gadget&amp;synd=open&amp;w=140&amp;h=75&amp;title=Alle+sider&amp;border=%23ffffff%7C3px%2C1px+solid+%23999999&amp;output=js"></script>
  </p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
        </tr>
        <tr bgcolor="#FAFAF4">
          <td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
          <td width="98%" class="sidebarHeader">&nbsp;&nbsp;NYHEDSMAIL</td>
        </tr>
        <tr>
          <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
        </tr>
        <tr>
          <td colspan="2" class="sidebarText"><script src="/shared/newsmail.js"></script>
              <p align="center"><a href="#" onClick="javascript:window.open('http://www.eco-info.dk/home/newsmail_okonet.asp','subwin','scrollbars=no,resizable=no,width=500,height=330,top=200,left=200');"><br>
                Modtag 
                nyhedsmail</a></p>
            <p align="center"><br>
              </p>
            <div align="center">
                <!--script type="text/javascript">
google_ad_client = "pub-1814652264778642";
/* 120x600, oprettet 04-03-08 */
google_ad_slot = "2048254750";
google_ad_width = 120;
google_ad_height = 600;

</script-->
                <!--script type="text/javascript" src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script-->
            </div></td>
        </tr>
      </table>
      </td>
  </tr>
</table>
<table width="90%"  border="0" align="center">
  <tr>
    <td height="30" class="sidebarHeader"><div align="center">
      <p><a href="sitemap.asp"> Sitemap<img src="/shared/graphics/layout/folder.gif" alt="Alle &Oslash;ko-net's sites" width="15" height="15" hspace="5" border="0"></a></p>
    </div></td>
  </tr>
</table></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="388" height="100%" valign="top" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> 
<!--
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="5" height="22"></td>
<td width="99%">Øko-info støttes af:</td>
<td><a href="about_4.asp"><img src="graphics/sponsorat_anim.gif" width="250" height="22" border="0"></a></td>
</tr>
<tr> 
<td colspan="3" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
-->
<table border="0" cellspacing="0" cellpadding="0"><tr><td valign="top">

</td></tr></table>
<!--vindue-->
<!--include virtual="log/ei/home/inc_leader.txt"-->
<!--Bæreklang billedskift-->
<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;SEKTIONER</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
<div align="center">
  <p>
  <!--#include file="writesections.asp"-->
    <!--include virtual="log/ei/home/inc_sections.txt"-->
    <br>
      </p>
  <p>&nbsp;</p>
  <p><a href="http://web.eco-net.dk/home/DetgodeProgram.asp" target="_self"></a><br>
    
      <br>
      </p>
	  <!--script type="text/javascript">
google_ad_client = "pub-1814652264778642";
/* 200x200, oprettet 22-02-08 */
google_ad_slot = "5847917114";
google_ad_width = 200;
google_ad_height = 200;

</script-->
<!--script type="text/javascript" src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script-->
</div></td>
</tr>
<tr> 
<td valign="bottom" align="left"> 
  <div align="center">
    <!--#include virtual="/shared/pagetools.asp"-->
  </div></td>
</tr>
</table>
</td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="180" valign="top" name="rightbar" bgcolor="#ECECDD"> 
<!--#include file="rightpane.asp"-->
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
    </tr>
    <tr bgcolor="#FAFAF4">
      <td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
      <td width="98%" align="left" valign="middle" class="sidebarHeader"><div align="left">&nbsp;&nbsp;KLIMAALARM</div></td>
    </tr>
    <tr>
      <td colspan="2" height="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
    </tr>
  </table>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td valign="top"><p align="center"> <a href="http://www.facebook.com/group.php?gid=42473770495&ref=mf#/group.php?gid=42473770495" target="_blank"><img src="/home/graphics/klimaalarm_cover.jpg" alt="Klimaalarm" width="154" height="150" vspace="5" border="0"></a><br>
        <br>
        <strong>NY st&oslash;tte-CD<br>
                  </strong>En musikalsk lyrik-samling, hvor den gr&oslash;nne tr&aring;d er klimakrisen, et problem der   ber&oslash;r os alle - ung som gammel. <br>
                  <br>
                  <strong>Klimaalarmen er g&aring;et!<br>
                  G&oslash;r noget!</strong> <br>
                  <br>
                  Bestil og l&aelig;s n&aelig;rmere om cd'en:<a href="http://www.klimaalarm.dk"><strong> www.klimaalarm.dk</strong></a><br>
                  <br>
      </p>
        </td>
    </tr>
  </table>

<!--#include file="cd_forside.asp"--></td>
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
<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-3636900-2";
urchinTracker();
</script>
</body>
</html>
<%
rsNews.Close()
%>

<!--include virtual="/shared/log.asp"-->
<%'reg("homeindex")%>
