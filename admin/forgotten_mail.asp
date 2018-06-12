<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsMail__MMColParam
rsMail__MMColParam = ""
if (Request("mailadr") <> "") then rsMail__MMColParam = Request("mailadr")
%>
<%

set rsMail = Server.CreateObject("ADODB.Recordset")
rsMail.ActiveConnection = MM_ecoinfo_STRING
rsMail.Source = "SELECT * FROM eisys_insiderusers WHERE email = '" + Replace(rsMail__MMColParam, "'", "''") + "';"

rsMail.CursorType = 0
rsMail.CursorLocation = 2
rsMail.LockType = 3
rsMail.Open()
rsMail_numRows = 0

%>
<% if rsMail.EOF then
answer="no-mail"
else
answer="mail"
%>
<%
Dim theTo,theFrom,theSubject,theBody
'theTo="web@eco-net.dk"
theTo = rsMail__MMColParam
theFrom="eco-net@eco-net.dk"
theSubject="Brugerdata til Øko-info"
theBody="Her er dine brugerdata til Øko-info: <br>" 

theBody=theBody & "<br>Brugernavn: " & rsMail.Fields.Item("username").Value 
theBody=theBody & "<br>Adgangskode: " & rsMail.Fields.Item("password").Value
theBody=theBody & "<br><br>Med Venlig hilsen Øko-net"
theBody=theBody & "<br><a href='http://www.eco-info.dk'>eco-info.dk</a>"
SendCDOMail theTo,theFrom,theSubject,theBody

end if
%>

<%
rsMail.Close()
%>
<% if rsOrg__MMColParam <>"" and rsOrg__MMColParam <> "0" then
rsOrg.Close()
end if
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
                	 <p align="center">&nbsp;</p>
              </div></td>
          </tr>
          <tr>
            <td scope="col"><div align="left">
                	 <p align="center"><span class="contentHeader1">
					 <% if answer="mail" then %>
					 Tak for interessen.<br />
					 Din adgangskode er på vej
					 <% end if %>
					 <% if answer="no-mail" then %>
					 Desv&aelig;rre kender vi ikke den emailadresse du angav.
					 <% end if %>
				     </span></p>
              </div></td>
          </tr>
        </table>
    <p align="center" class="contentHeader2"><a href="/home/insider.htm">retur</a></p>
    <p align="center" class="contentHeader2"><a href="/home/index.asp">retur til &Oslash;ko-info</a> </p></td>
    <td valign="top" bgcolor="#E7F1F7"><p align="center" class="contentHeader1">&nbsp;</p>
        <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><%if answer="no-mail" then %>
			  <br />
			  Hvis du eller din organisation er oprettet i &Oslash;ko-info,<br />
			kan det skyldes at emailadrseen er &aelig;ndret.<br />
			  Ring eller mail venligst til os.
			  <p><a href="/home/insider.htm">Hvis du ikke er oprettet i &Oslash;ko-info,<br />
			  kan du g&oslash;re det her.</a></p>
			<%end if%>
            <p>MVH &Oslash;ko-net, <a href="mailto:kmh@eco-net.dk">kmh@eco-net.dk</a>, 62 24 43 24 </p>
			
			<%if app="edituser" then %>
			<p>Vi retter snarest dine data i databasen.</p>
            <p>Som regel n&aring;r vi det samme dag,<br />
              eller senest n&aelig;ste arbejdsdag.
</p>
            <p>MVH &Oslash;ko-net, <a href="mailto:kmh@eco-net.dk">kmh@eco-net.dk</a>, 62 24 43 24 </p>
			<%end if%>
            <p>&nbsp;</p></td>
          </tr>
        </table>
    </td>
  </tr>
</table>
</body>
</html>
