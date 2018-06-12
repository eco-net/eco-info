<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include file="act_sqlcalc.asp" -->
<!--#include virtual="/shared/showimage.asp" -->
<!--#include file="nav.asp"-->
<!--#include virtual="shared/common.asp" -->
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM eiart_maindata m  WHERE 0=0  "
rsPageData.Source=replace(rsPageData.Source,"0=0",theWhere)
'rsPageData.Source=replace(rsPageData.Source,"eiart_maindata m",theFrom)

rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<%
Dim rsImg3__theID
rsImg3__theID = "0"
if (request("id") <> "") then rsImg3__theID = request("id")
%>
<%
set rsImg3 = Server.CreateObject("ADODB.Recordset")
rsImg3.ActiveConnection = MM_ecoinfo_STRING
rsImg3.Source = "SELECT i.id, i.subtext,i.source  FROM eiart_r_img r LEFT JOIN images i ON r.imgid=i.id  WHERE artid=" + Replace(rsImg3__theID, "'", "''") + " AND imgsize='3'  ORDER BY sort"
rsImg3.CursorType = 0
rsImg3.CursorLocation = 2
rsImg3.LockType = 3
rsImg3.Open()
rsImg3_numRows = 0
%>
<%
Dim rsImgR__theID
rsImgR__theID = "0"
if (request("id") <> "") then rsImgR__theID = request("id")
%>
<%
Dim Recordset1__theID
Recordset1__theID = "0"
if (request("id") <> "") then Recordset1__theID = request("id")
%>
<%
set rsImgR = Server.CreateObject("ADODB.Recordset")
rsImgR.ActiveConnection = MM_ecoinfo_STRING
rsImgR.Source = "SELECT i.id, i.subtext,i.source  FROM eiart_r_img r LEFT JOIN images i ON r.imgid=i.id  WHERE artid=" + Replace(rsImgR__theID, "'", "''") + " AND imgsize='R'  ORDER BY sort"
rsImgR.CursorType = 0
rsImgR.CursorLocation = 2
rsImgR.LockType = 3
rsImgR.Open()
rsImgR_numRows = 0
%>
<%
Dim rsCat__theID
rsCat__theID = "0"
if (request("id") <> "") then rsCat__theID = request("id")
%>
<%
set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.ActiveConnection = MM_ecoinfo_STRING
rsCat.Source = "SELECT cm.name,cm.id  FROM eiart_cat_maindata cm LEFT JOIN eiart_r_cat rc ON cm.id=rc.catid  WHERE rc.artid=" + Replace(rsCat__theID, "'", "''") + ""
rsCat.CursorType = 0
rsCat.CursorLocation = 2
rsCat.LockType = 3
rsCat.Open()
rsCat_numRows = 0
%>
<%
Dim rsSbj__theID
rsSbj__theID = "0"
if (request("id") <> "") then rsSbj__theID = request("id")
%>
<%
set rsSbj = Server.CreateObject("ADODB.Recordset")
rsSbj.ActiveConnection = MM_ecoinfo_STRING
rsSbj.Source = "SELECT sm.id, sm.name  FROM eisbj_maindata sm LEFT JOIN eiart_r_sbj rs ON sm.id=rs.sbjid WHERE rs.artid=" + Replace(rsSbj__theID, "'", "''") + ""
rsSbj.CursorType = 0
rsSbj.CursorLocation = 2
rsSbj.LockType = 3
rsSbj.Open()
rsSbj_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsImg3_numRows = rsImg3_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
rsImgR_numRows = rsImgR_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Repeat3__numRows = -1
Dim Repeat3__index
Repeat3__index = 0
rsSbj_numRows = rsSbj_numRows + Repeat3__numRows
%>
<!--include virtual="/shared/inc_detailnavigation.asp" -->
<!--include file="act_detail.asp" -->
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>&Oslash;ko-info - Artikler om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
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
<!--#include virtual="/art/inc_search.asp"-->
<%dim editpage,editid
editpage="art"
editid=OrgFromArt(rsPageData("id"))
%>
<!--#include virtual="/admin/inc_mypages.asp"-->
<form name="form1" method="post" action="/shared/print_art_dk.asp" target="_blank">
<div align="center">Artiklen kan udskrives alene<br>
<br>
<input type="submit" name="Submit" value="Udskriv" class="formSubmitbutton">
<input type="hidden" name="id" value="<%=request("id")%>">
</div>
</form>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" --> 
<table width="100%" border="0" cellspacing="0" cellpadding="16" background="/dgb/graphics/detail_header_dgb_bkgrd.gif" name="detailHeader">
<tr> 
<td valign="top"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif" name="detailContents">
<tr> 
<td colspan="2" class="contentHeader1" align="left"><%=makeRed(toString(rsPageData.Fields.Item("title").Value))%>&nbsp;&nbsp;&nbsp; 
<%IF rsPageData.Fields.Item("subtitle").Value<>"" THEN response.write "<br><span class='listheader'> " & rsPageData.Fields.Item("subtitle").Value & "</span>"%>
<%IF LEN(request.cookies("eiuserid"))>0 THEN%>
<a href="http://insider.eco-info.dk/art/edit.asp?id=<%=(rsPageData.Fields.Item("id").Value)%>" class="sitetag" target="_blank">Rediger 
kortet</a> 
<%END IF%>
</td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td height"20">&nbsp;</td>
<td height="20" align="right" nowrap></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="1" height="5"></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top" width="364"> 
<table border="0" cellspacing="0" cellpadding="0" width="100%">
<tr> 
<td class="contentHeader2"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><a href="list.asp">Tilbage til oversigt</a></td>
<td> 
<div align="right"><%=navHTML()%></div>
</td>
</tr>
</table>
<p>&nbsp;</p>
</td>
</tr>
<tr> 
<td valign="top" width="95%"> 
<p><%=makeRed(toString(rsPageData.Fields.Item("descr").Value))%></p>
<p align="center">&nbsp; 
<% If Not rsImg3.EOF Or Not rsImg3.BOF Then %>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsImg3.EOF)) 
%>
<img src="<%=getImage(rsImg3.Fields.Item("id").Value,"3")%>"> 
<br>
<%=(rsImg3.Fields.Item("subtext").Value)%><br>
<i><%=(rsImg3.Fields.Item("source").Value)%></i><br>
<br>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsImg3.MoveNext()
Wend
%>
<% End If ' end Not rsImg3.EOF Or NOT rsImg3.BOF %>
</p>
<p><span class="contentHeader2">Andre oplysninger</span><br>
Findes i kategorien: <a href="list.asp?artcat=<%=(rsCat.Fields.Item("id").Value)%>"><%=(rsCat.Fields.Item("name").Value)%></a><br>
Under emneord: 
<% 
While ((Repeat3__numRows <> 0) AND (NOT rsSbj.EOF)) 
%>
<a href="list.asp?artsbj=<%=(rsSbj.Fields.Item("id").Value)%>"><%=(rsSbj.Fields.Item("name").Value)%></a>, 
<% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  rsSbj.MoveNext()
Wend
%>
</p>
</td>
</tr>
</table>
</td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td valign="top"> 
<table border="0" cellspacing="0" cellpadding="0" width="160">
<tr> 
<td class="contentHeader2" width="160">Fakta</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td width="180"> 
<p> 
<% if rsPageData.Fields.Item("author").Value <> "" then %>
<br>
<span class="faktaboksheader">Forfatter</span><br>
<%=(rsPageData.Fields.Item("author").Value)%> <br>
<% end if %>
<% if rsPageData.Fields.Item("lang").Value <> "" then %>
<br>
<span class="faktaboksheader">Sprog</span><br>
<%=(rsPageData.Fields.Item("lang").Value)%> <br>
<% end if %>
<br>
<span class="faktaboksheader">Udgiver</span><br>
<%=replace(rsPageData.Fields.Item("editor").Value,chr(13),"<br>")%> <br>
<% if rsPageData.Fields.Item("authordate").Value <> "" then %>
<br>
<span class="faktaboksheader">Udgivelses&aring;r</span><br>
<%=(rsPageData.Fields.Item("authordate").Value)%> <br>
<% end if %>
<% if rsPageData.Fields.Item("editormail").Value <> "" then %>
<br>
<span class="faktaboksheader">Udgivers mail</span><br>
<a href="mailto:<%=(rsPageData.Fields.Item("editormail").Value)%>"><%=(rsPageData.Fields.Item("editormail").Value)%></a> <br>
<% end if %>
<% if rsPageData.Fields.Item("editorwww").Value <> "" then %>
<br>
<span class="faktaboksheader">Udgivers WWW</span><br>
<a href="http://<%=(rsPageData.Fields.Item("editorwww").Value)%>" title="Åbner i et nyt vindue" target="_blank"><%=(rsPageData.Fields.Item("editorwww").Value)%></a> <br>
<% end if %>
<% if rsPageData.Fields.Item("webaddress").Value <> "" then %>
<br>
<span class="faktaboksheader">Publikationen kan ses p&aring;:</span><br>
<a href="http://<%=(rsPageData.Fields.Item("webaddress").Value)%>" title="Åbner i et nyt vindue" target="_blank"><%=(rsPageData.Fields.Item("webaddress").Value)%></a> <br>
<% end if %>
</p>
<% If Not rsImgR.EOF Or Not rsImgR.BOF Then %>
<p align="center"> 
<% 
While ((Repeat2__numRows <> 0) AND (NOT rsImgR.EOF)) 
%>
<img src="<%=getImage(rsImgR.Fields.Item("id").Value,"R")%>"> 
<br>
<%=(rsImgR.Fields.Item("subtext").Value)%><br>
<i><%=(rsImgR.Fields.Item("source").Value)%></i> <br>
<br>
<% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsImgR.MoveNext()
Wend
%>
<br>
</p>
<% End If ' end Not rsImgR.EOF Or NOT rsImgR.BOF %>
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
<%
rsImg3.Close()
%>
<%
rsImgR.Close()
%>
<%
rsCat.Close()
%>
<%
rsSbj.Close()
%>
<!--include virtual="/shared/log.asp"-->

