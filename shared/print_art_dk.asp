<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsPageData__theID
rsPageData__theID = "0"
if (request("id") <> "") then rsPageData__theID = request("id")
%>
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM eiart_maindata  WHERE id=" + Replace(rsPageData__theID, "'", "''") + ""
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
<html>
<head>
<title><%=(rsPageData.Fields.Item("title").Value)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="650" border="0" cellspacing="5" cellpadding="5" align="center">
<tr> 
<td colspan="2"><img src="/shared/graphics/logo.gif" width="180" height="33"><br>
<span class="sitetag">www.eco-info.dk - &Oslash;ko-net, Svendborgvej 9 - DK 5762 
V.Skerninge - (+45) 62 24 43 24 - eco-net@eco-net.dk</span><br>
</td>
</tr>
<tr> 
<td colspan="2">&nbsp;</td>
</tr>
<tr> 
<td valign="top" width="480"> <span class="contentHeader1"><%=(rsPageData.Fields.Item("title").Value)%><br>
</span> 
<%if(rsPageData.Fields.Item("subtitle").Value)<>"" then %>
<span class="listheader"><%=(rsPageData.Fields.Item("subtitle").Value)%></span> 
<% end if %>
</td>
<td valign="top" class="listheader" width="160"><br>
<span class="sidebarHeader"><%=Date%></span></td>
</tr>
<tr> 
<td valign="top" width="480"> 
<p><br>
<%=(rsPageData.Fields.Item("descr").Value)%></p>
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
</td>
<td valign="top" nowrap width="160"> 
<p> <span class="listheader"><br>
Forfatter</span><br>
<%if(rsPageData.Fields.Item("author").Value)<>"" then %>
<%=(rsPageData.Fields.Item("author").Value)%><br>
<% end if %>
<%if(rsPageData.Fields.Item("authordescr").Value)<>"" then %>
<%=(rsPageData.Fields.Item("authordescr").Value)%><br>
<% end if %>
<%if(rsPageData.Fields.Item("authordate").Value)<>"" then %>
<%=(rsPageData.Fields.Item("authordate").Value)%><br>
<% end if %>
<%if(rsPageData.Fields.Item("webaddress").Value)<>"" then %>
<%=(rsPageData.Fields.Item("webaddress").Value)%> 
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
<span class="listheader">Udgiver</span><br>
<br>
<%if(rsPageData.Fields.Item("editor").Value)<>"" then %>
<%=(rsPageData.Fields.Item("editor").Value)%><br>
<% end if %>
<%if(rsPageData.Fields.Item("editoraddress").Value)<>"" then %>
<%=(rsPageData.Fields.Item("editoraddress").Value)%> <br>
<% end if %>
<%if(rsPageData.Fields.Item("editormail").Value)<>"" then %>
<%=(rsPageData.Fields.Item("editormail").Value)%><br>
<% end if %>
<%if(rsPageData.Fields.Item("editorwww").Value)<>"" then %>
<%=(rsPageData.Fields.Item("editorwww").Value)%> 
<% end if %>
</td>
</tr>
</table>
</body>
</html>
<%
rsPageData.Close()
%>
<%
rsImg3.Close()
%>
<%
rsImgR.Close()
%>
<script language="javascript">
print();
</script>


