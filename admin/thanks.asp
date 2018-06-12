<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
name=request("name")
app=request("app")
title=request("title")
appname=""
mailbody=""
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Tak</title>
<link href="/shared/styles.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="750" border="1" align="center" cellpadding="5" cellspacing="0" bordercolor="#003300" bgcolor="#FFFFFF">
  <tr>
    <th colspan="2" bgcolor="#1979B5" scope="col"><div align="left"><img src="/shared/graphics/logo.gif" width="180" height="33" /></div></th>
  </tr>
  <tr>
    <td valign="top" bgcolor="#E7F1F7"><p align="center" class="contentHeader1">&nbsp;</p>
        <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr>
            <td scope="col"><div align="left">
                	 <p align="center"><span class="contentHeader1">
					 <%=name%>
					  </span></p>
              </div></td>
          </tr>
          <tr>
            <td scope="col"><div align="left">
                	 <p align="center"><span class="contentHeader1">
					 <% if app="newuser" then %>
					 Tak for oprettelsen i &Oslash;ko-info
					 <% end if %>
					 <% if app="edituser" then %>
					 Tak for rettelsen i &Oslash;ko-info
					 <% end if %>
					  <% if app="newarr" then %>
					 Tak for indsendelsen af arrangementet til &Oslash;ko-info
					 <% end if %>
					 <% if app="editarr" then %>
					 Tak for rettelsen af arrangementet til &Oslash;ko-info
					 <% end if %>
					 <% if app="newpub" then 
					  appname="Publikation"
					 mailbody="Til &Oslash;ko-net. - De vedh&aelig;ftede billeder h&oslash;rer til " & appname & " med titlen: " & title & ". - MVH - " & name
					 %>
					 Tak for indesendelsen af publikationen til &Oslash;ko-info
					 <% end if %>
					 <% if app="editpub" then 
					  appname="Publikation"
					 mailbody="Til &Oslash;ko-net. - De vedh&aelig;ftede billeder h&oslash;rer til " & appname & " med titlen: " & title & ". - MVH - " & name
					 %>
					 Tak for rettelsen af publikationen til &Oslash;ko-info
					 <% end if %>
					 <% if app="newart" then
					 appname="Artikel"
					 mailbody="Til &Oslash;ko-net. - De vedh&aelig;ftede billeder h&oslash;rer til " & appname & " med titlen: " & title & ". - MVH - " & name
					  %>
					 Tak for indsendelsen af artiklen til &Oslash;ko-info
					 <% end if %>
					 <% if app="editart" then 
					 appname="Artikel"
					 mailbody="Til &Oslash;ko-net. - De vedh&aelig;ftede billeder h&oslash;rer til " & appname & " med titlen: " & title & ". - MVH - " & name
					 %>
					 Tak for rettelsen af artiklen til &Oslash;ko-info
					 <% end if %>
					 <% if app="newopsl" then 
					 appname="Opslag"
					 mailbody="Til &Oslash;ko-net. - De vedh&aelig;ftede billeder h&oslash;rer til " & appname & " med titlen: " & title & ". - MVH - " & name
					 %>
					 Tak for opslaget til &Oslash;ko-info
					 <% end if %>
				     </span></p>
              </div></td>
          </tr>
        </table>
    <% if mailbody<>"" then %>
	<p align="center" class="contentHeader2">
	<a href="mailto:kmh@eco-net.dk?subject=Billeder fra <%=name%>&Body=<%=mailbody%>">Indsend Billeder</a></p>
	<%end if%>
    <p align="center" class="contentHeader2"><a href="/home/index.asp">Retur til &Oslash;ko-info</a> </p></td>
    <td valign="top" bgcolor="#E7F1F7"><p align="center" class="contentHeader1">&nbsp;</p>
        <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><%if app="newuser" then %><p>Vi opretter snarest dine data i databasen.</p>
            <p>Du vil modtage en email, n&aring;r du er blevet oprettet  p&aring; &Oslash;ko-info,<br />
            med brugernavn og adgangskode,<br />
              s&aring; du senere kan logge ind,<br />
              og 
              oprette dine arrangementer og publikationer i databasen.</p>
              <%end if%>

            <%if app="edituser" then %>
			<p>Vi retter snarest dine data i databasen.</p>
            <%end if%>
			
            <%if app="newarr" then %>
            
            <p>Vi opretter snarest dit arrangement i databasen.</p>
            <%end if%>
              
             <%if app="editarr" then %>
            <p>Vi retter snarest dit arrangement i databasen.</p>
            <%end if%>
			
            <%if app="newpub" then %>
            <p>Vi opretter snarest din publikation i databasen.</p>
            <%end if%>
            
			<%if app="editpub" then %>
            <p>Vi retter snarest din publikation i databasen.</p>
            <%end if%>
			
			<%if app="newart" then %>
            <p>Vi opretter snarest din artikel i databasen.</p>
            <%end if%>
            
			<%if app="editart" then %>
            <p>Vi retter snarest din artikel i databasen.</p>
            <%end if%>
			
			<%if app="newopsl" then %>
            <p>Vi opretter snarest din opslag i databasen.</p>
            <%end if%>
			<p>Som regel n&aring;r vi det samme dag,<br />
              eller senest n&aelig;ste arbejdsdag.</p>
<p>MVH &Oslash;ko-net, <a href="mailto:kmh@eco-net.dk">kmh@eco-net.dk</a>, 62 24 43 24            </p>
            <p>&nbsp;</p></td>
          </tr>
        </table>
    </td>
  </tr>
</table>
</body>
</html>
