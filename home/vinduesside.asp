<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Connections/ecoinfo.asp" -->

<%
set rsChosen = Server.CreateObject("ADODB.Recordset")
rsChosen.ActiveConnection = MM_ecoinfo_STRING
rsChosen.Source = "SELECT id FROM eivindue_maindata  WHERE chosen=1"
if request("id")<>"" then 
theWhere="id=" & request("id")
replace rsChosen.Source,"chosen=1",theWhere
end if
rsChosen.CursorType = 0
rsChosen.CursorLocation = 2
rsChosen.LockType = 3
rsChosen.Open()
page_id=rsChosen("id")
rsChosen.Close()
%>

<%
set rsVin = Server.CreateObject("ADODB.Recordset")
rsVin.ActiveConnection = MM_ecoinfo_STRING
rsVin.CursorType = 0
rsVin.CursorLocation = 2
rsVin.LockType = 3
rsVin.Source = "SELECT * FROM eivindue_item WHERE page_id=" & page_id & " ORDER BY thesort"
rsVin.Open()
%>
<%
set rsVins = Server.CreateObject("ADODB.Recordset")
rsVins.ActiveConnection = MM_ecoinfo_STRING
rsVins.CursorType = 0
rsVins.CursorLocation = 2
rsVins.LockType = 3
rsVins.Source = "SELECT id,title FROM eivindue_maindata WHERE chosen=0 ORDER BY title"
if request("id")<>"" then 
replace rsVins.Source,"chosen=0",""
end if 
rsVins.Open()
%>

<html>
<head>
<title>&Oslash;ko-info - guide til alt om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
<meta name="keywords" content="økologi, miljø, bæredygtig udvikling, natur, grøn, landbrug, fødevarer, agenda 21, energi, byøkologi, forurening, økologisk, ecology, bæredygtighed, agenda21, besøgslandbrug, gårdbutik, grøn guide, naturvejleder, genbrug, klima, økosamfund, trafik, økologisk byggeri, sundhed, helse, helsekost, permakultur, alternativ, vand, atomkraft">
<meta name="description" content="Øko-info er Danmarks økologiske portal, med alt om økologi, miljø og bæredygtig udvikling.">
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
</tr></table><!-- #EndLibraryItem -->
<!--#include virtual="/shared/showimage.asp" -->
<table width="750" border="0" cellspacing="0" cellpadding="10" name="Contentarea">
<tr>
  <td width="180" valign="top" bgcolor="#DAD8D6">
	<%  if not rsVin.EOF then
	  while not rsVin.EOF %>
    <% if rsVin("col")=1 then %>
    <% if rsVin("img_id")=0 then %>
    <%=rsVin("tekst")%><br><br>
	<% if rsVin("source")<>"" then %>
				<i>Kilde: <%=rsVin("source")%></i><br><br>
			<% end if %>
			<% if rsVin("link")<>"" then %>
				<a href="<%=rsVin("link")%>" target="_blank"><%=rsVin("linktext")%></a><br>
			<% end if %>

    <% else %>
    <div align="center"> <img src="<%=getImage(rsVin("img_id"),"R")%>"> <br>
        <% =rsVin("img_text") %>
        <br>
    </div>
    <% end if %>
    <br>
    <% end if 
	%>
	<% rsVin.MoveNext
	  wend
	  rsVin.Close()
	   end if %> 
	   <% if not rsVins.EOF then %><br><br>
	   <table width="95%"  border="0" align="center">
      <tr>
        <td bgcolor="#FFFFCC">
		 <%
		response.write "<b>L&aelig;s ogs&aring;:<b><br><br>"
		while not rsVins.EOF
		response.write "<a href='vinduesside.asp?id=" & rsVins("id") & "'>" & rsVins("title") & "</a><br>"
		rsVins.MoveNext
		wend
		%>
		
 </td>
      </tr>
    </table> 
	<%
		end if 
		rsVins.Close()
	%>	</td> 
  
<td valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="5">
    
    <tr>
      <td>
	  <% rsVin.Open() 
		if not rsVin.EOF then
	  while not rsVin.EOF
	  %>
    <% if rsVin("col")=2 then %>
        <% if rsVin("img_id")=0 then %>
         	<%=rsVin("tekst")%><br><br>
			<% if rsVin("source")<>"" then %>
				<i>Kilde: <%=rsVin("source")%></i><br><br>
			<% end if %>
			<% if rsVin("link")<>"" then %>
				<a href="<%=rsVin("link")%>" target="_blank"><%=rsVin("linktext")%></a><br>
			<% end if %>
        <% else %>
        	<div align="center"> <img src="<%=getImage(rsVin("img_id"),"3")%>"><br>
            <% =rsVin("img_text") %>
        	</div>
        <% end if %>
        
     <% end if %>
      </td>
    </tr>
    <% rsVin.MoveNext
	  wend
	  rsVin.Close()
	   end if %>
</table></td>

<td width="180" valign="top" bgcolor="#DAD8D6" name="rightbar"> 
  <table width="180" border="0" cellspacing="0" cellpadding="0">
<tr>
  <td valign="top" bgcolor="#DAD8D6"><% rsVin.Open()
		if not rsVin.EOF then
	  while not rsVin.EOF
	  %>
          <% if rsVin("col")=3 then %>
          <% if rsVin("img_id")=0 then %>
           <%=rsVin("tekst")%><br><br>
	<% if rsVin("source")<>"" then %>
				<i>Kilde: <%=rsVin("source")%></i><br><br>
			<% end if %>
			<% if rsVin("link")<>"" then %>
				<a href="<%=rsVin("link")%>" target="_blank"><%=rsVin("linktext")%></a><br>
			<% end if %>

          <% else %>
          <div align="center"> <img src="<%=getImage(rsVin("img_id"),"R")%>">
              <% =rsVin("img_text") %>
              <br>
          </div>
          <% end if %>
         <br><br>
          <% end if %>
        
      <% rsVin.MoveNext
	  wend
	  rsVin.Close()
	   end if %></td> 
<td colspan="2" bgcolor="#DAD8D6" class="sidebarText"> 
<!--#include virtual="log/ei/tema/inc_annoncer.asp"--></td>
</tr>
</table></td>
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

<!--#include virtual="/shared/log.asp"-->
<%reg("temaindex")%>
