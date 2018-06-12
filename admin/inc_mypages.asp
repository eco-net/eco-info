<% if Session("orgid")<>"" then %>
<br />
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td bgcolor="#1979B5" class="contentHeader1">Mine sider </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#E7F1F7"><p class="listheader"> <a href="/admin/new_arr.asp">Nyt i &Oslash;ko-Kalenderen </a><br />
        <a href="/admin/new_pub.asp">Nyt i Det Gr&oslash;nne Bibliotek </a> <br />
        <a href="/admin/new_art.asp">Ny Artikel</a></p>
        <p class="listheader"><a href="/dgs/detail.asp?id=<%=Session("orgid")%>">Stamdata</a><br />
            <%if Session("cal")<>"0" then%>
            <a href="/ok/list.asp?orgid=<%=Session("orgid")%>">I &Oslash;ko-Kalenderen </a>
            <%end if%>
          &nbsp;<br />
            <%if Session("lib")<>"0" then%>
            <a href="/dgb/list.asp?orgid=<%=Session("orgid")%>">I Det Gr&oslash;nne Bibliotek </a>
            <%end if%>
          &nbsp;<br /><%if Session("art")<>"0" then%>
            <a href="/art/list.asp?orgid=<%=Session("orgid")%>">Artikler</a>
            <%end if%>
          &nbsp;</p>
        <p class="listheader">
		<% if editpage="s" and editid=Session("orgid") then%>
		<a href="/admin/edit_user.asp?id=<%=Session("orgid")%>">Ret Stamdata</a>
		<%end if%>
		<% if editpage="a" and editid=Session("orgid") then%>
		<a href="/admin/edit_arr.asp?id=<%=rsPageData("id")%>">Ret Arrangement</a>
		<%end if%>
		<% if editpage="p" and editid=Session("orgid") then%>
		<a href="/admin/edit_pub.asp?id=<%=rsPageData("id")%>">Ret Publikation</a>
		<%end if%>
		<% if editpage="art" and editid=Session("orgid") then%>
		<a href="/admin/edit_art.asp?id=<%=rsPageData("id")%>">Ret Artikel</a>
		<%end if%>
		</p></td>
  </tr>
</table>
<%else%>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td bgcolor="#1979B5" class="contentHeader1">Mine sider </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#E7F1F7"><p class="listheader">Gr&oslash;nne organisationer - NGO'ere som offentlige - og private personer kan oprette   og pr&aelig;sentere sig selv i databasen
        <!-- - med mulighed for at tilknytte   arrangementer til &Oslash;ko-kalenderen og publikationer i Det Gr&oslash;nne Bibliotek -->
        </p>
      <p class="listheader">
        <input type="button" name="Button" value="Kom med / Log ind" onClick="javascript:location.href='/home/insider.htm'"/>
        <br />
      </p></td>
  </tr>
</table>
<% end if %>

