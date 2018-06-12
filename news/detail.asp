<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<!--#include file="act_sqlcalc.asp" -->
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

Dim rsPageData__theID
rsPageData__theID = "0"
if (request("id") <> "") then rsPageData__theID = request("id")
%>
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM bu_Artikel INNER JOIN bu_Artikel_site ON bu_Artikel.Artikel_ID=bu_Artikel_site.artikel_id  WHERE bu_Artikel.Artikel_ID=" + Replace(rsPageData__theID, "'", "''") + ";"
'IF theWhere<>"" THEN rsPageData.Source=replace(rsPageData.Source,"0=0",theWhere)
'IF theFrom<>"" THEN rsPageData.Source=replace(rsPageData.Source,"FROM eilib_sitedata s LEFT JOIN eilib_maindata m ON s.libid=m.id",theFrom)
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>

<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>&Oslash;ko-info - Aktuelt om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
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
<!-- #BeginEditable "menu" --> 
<%DIM curtab
curtab=5%>
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
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td class="sidebarHeader">
<div align="center"><a href="/news/index.asp">Nyhedsoversigt</a></div>
</td>
</tr>
</table>
<p>&nbsp;</p>
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" --> 
<table width="100%" border="0" cellspacing="0" cellpadding="16" background="/news/graphics/detail_header_news_bkgrd.gif" name="detailHeader">
<tr> 
<td valign="top"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif" name="detailContents">
<tr> 
<td class="contentHeader1" colspan="2"><span class="contentHeader1"><%=makeRed(toString(rsPageData.Fields.Item("Header").Value))%></span><img src="/shared/graphics/spacer.gif" width="3" height="20"></td>
</tr>
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="5"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td height="20"><a href="javascript:history.go(-1)">tilbage</a></td>
<td height="20" align="right" nowrap>&nbsp;</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif" name="detailContents">

<tr> 
<td valign="top" width="400"> 
<p>&nbsp;<%=makeRed(toString(rsPageData.Fields.Item("Content").Value))%></p>
<div align="center">
  <% if Len(rsPageData.Fields.Item("imagefilename2").Value)>0 then %>
  <br>
    <img src="<%=picasaImgSize(rsPageData("imagefilename2"),"/s400")%>" width=360/><br>
    <%=rsPageData("img_text2")%><br>
    <% end if %>
</div></td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="10" height="3"></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="10" height="3"></td>
<td valign="top"> 
<table border="0" cellspacing="0" cellpadding="0">
<tr> 
<td nowrap >oprettet 
<%theDate=(rsPageData.Fields.Item("create_date").Value)
dato=datepart("d",theDate) & "-" & datepart("m",theDate) & "-" & datepart("yyyy",theDate)%>
<%=dato%><br>
<br>
</td>
</tr>
<tr> 
<td nowrap class="contentHeader2">L&aelig;s mere</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>


<tr> 
<td nowrap><br>
 
<%if(rsPageData.Fields.Item("FileName").Value)<>"" then %>
<p><span class="faktaboksheader"> <a href="<%=(rsPageData.Fields.Item("FileName").Value)%>" title="<%=(rsPageData.Fields.Item("FileName").Value)%>" <%if rsPageData.Fields.Item("newwin").Value =0 then response.write("target='_blank'")%>>
<% if rsPageData.Fields.Item("linktext").Value<>"" then %>
<%=(rsPageData.Fields.Item("linktext").Value)%>
<%else %>
L&aelig;s mere..
<%end if %></a> </span><br>
</p>
<%end if %>
<%if(rsPageData.Fields.Item("FileName2").Value)<>"" then %>
<p><span class="faktaboksheader"> <a href="<%=(rsPageData.Fields.Item("FileName2").Value)%>" title="<%=(rsPageData.Fields.Item("FileName2").Value)%>" <%if rsPageData.Fields.Item("newwin").Value =0 then response.write("target='_blank'")%>>
<% if rsPageData.Fields.Item("linktext2").Value<>"" then %>
<%=(rsPageData.Fields.Item("linktext2").Value)%>
<%else %>
L&aelig;s mere..
<%end if %></a> </span><br>
</p>
<%end if %>
<%if(rsPageData.Fields.Item("FileName3").Value)<>"" then %>
<p><span class="faktaboksheader"> <a href="<%=(rsPageData.Fields.Item("FileName3").Value)%>" title="<%=(rsPageData.Fields.Item("FileName3").Value)%>" <%if rsPageData.Fields.Item("newwin").Value =0 then response.write("target='_blank'")%>>
<% if rsPageData.Fields.Item("linktext3").Value<>"" then %>
<%=(rsPageData.Fields.Item("linktext3").Value)%>
<%else %>
L&aelig;s mere..
<%end if %></a> </span><br>
</p>
<%end if %>
<%if(rsPageData.Fields.Item("link4").Value)<>"" then %>
<p><span class="faktaboksheader"> <a href="<%=(rsPageData.Fields.Item("link4").Value)%>" title="<%=(rsPageData.Fields.Item("link4").Value)%>" <%if rsPageData.Fields.Item("newwin").Value =0 then response.write("target='_blank'")%>>
<% if rsPageData.Fields.Item("linktext4").Value<>"" then %>
<%=(rsPageData.Fields.Item("linktext4").Value)%>
<%else %>
L&aelig;s mere..
<%end if %></a> </span><br>
</p>
<%end if %>
</td>
</tr>

<tr> 
<td align="center"> 
  <p><br>
      <!--#include virtual="/shared/showimage.asp" -->
      <% if (rsPageData.Fields.Item("img_id").Value)>0 then %>
      <img src="<%=getImage(rsPageData.Fields.Item("img_id").Value,"R")%>"><br>
      <%=rsPageData("img_text")%><br>
    <br>
    
      <% end if %>
      <% if (rsPageData.Fields.Item("img_id2").Value)>0 then %>
      <img src="<%=getImage(rsPageData.Fields.Item("img_id2").Value,"R")%>"><br>
      <%=rsPageData("img_text2")%><br>
    
      <% end if %>
      <% if Len(rsPageData.Fields.Item("imagefilename1").Value)>0 then %>
    <br>
      <img src="<%=picasaImgSize(rsPageData("imagefilename1"),"/s144")%>" /><br>
      <%=rsPageData("img_text")%><br>
      <% end if %>
</p>
  <p>&nbsp;  </p></td>
</tr>

</table>

<br>
<table width="100%" border="0" cellpadding="5">
  <tr>
    <td>
	<%if rsPageData("miljoinfo")>0 then%>
	<hr>    
      Denne nyhed er fra MILJ&Oslash; INFO<br>
      MILJ&Oslash; INFO er en elektronisk informationsservice til alle, der besk&aelig;ftiger sig professionelt med milj&oslash; og arbejdsmilj&oslash;.<br>
      F&aring; mere info p&aring;: <a href="http://www.miljoinfo.dk" target="_blank">www.miljoinfo.dk</a><br>
      &Oslash;ko-net bringer dagligt en til to nyheder fra MILJ&Oslash; INFO.
	  <hr>
	  <% end if %></td>
  </tr>
</table></td>
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
