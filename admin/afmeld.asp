<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<%
id=request("id")
name=request("name")
firstname=request("firstname")
lastname=request("lastname")
stopped=0
remove=0
why=""
if request("stopped")<>"" then
stopped=1
why=why & " oph�rt "
end if
if request("remove")<>"" then
remove=1
why=why & ", fjernes "
end if
if request("update")<>"" then
style="<style type='text/css'><!--Body,Table {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px;}--></style>"
userinput=style & "<table width=400 border=1 cellspacing=0 cellpadding=5><tr><td><b>Afmelding</b></td><td><b>Eco-info</b></td></tr><tr><td>Navn</td><td>" & name & "</td></tr><tr><td>Fornavn</td><td>" & firstname & "&nbsp;</td></tr><tr><td>Efternavn</td><td>" & lastname & "&nbsp;</td></tr><tr><td>�rsag</td><td>" & why & "&nbsp;</td><tr><tr><td>Besked</td><td>" & request("besked") & "&nbsp;</td></tr><tr><td>I Insider</td><td><a href='http://insider.eco-info.dk/dgs/edit.asp?id=" & id & "'>Link</a>&nbsp;</td></tr></table>"
theBody="Hej Kirsten<br>En bruger �nsker at blive afmeldt p� �ko-info " & Time() & "<br>Vil du rette ham/hende i Insider/Filemaker. Tak" & userinput
'response.write theBody
'response.end
SendCDOMail "kmh@eco-net.dk","web@eco-net.dk","AFMELDING P� �KO-INFO",theBody
SendCDOMail "eco-net@eco-net.dk","web@eco-net.dk","AFMELDING P� �KO-INFO",theBody
response.redirect "edit_user.asp?id=" & id
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
                	 <p align="center" class="contentHeader1">Afmelding</p>
                	 <p align="center">Vi afmelder din organisation i &Oslash;ko-info<br />
                	 med bem. <%=why%><br />
           	  </p>
                	 <form id="form1" name="form1" method="post" action="">
                	   <div align="center">
                	     <p>Evt. besked til &Oslash;ko-net:
                	       <br />
                	       <textarea name="besked" id="besked"></textarea>
                	     </p>
                	     <p>
                	       <input type="submit" name="Submit" value="OK" />
               	           <input type="button" name="Submit2" value="Fortryd" onclick="javascript:history.go(-1);"/>
               	           <input name="id" type="hidden" id="id" value="<%=id%>" />
               	           <input name="name" type="hidden" id="name" value="<%=name%>" />
               	           <input name="firstname" type="hidden" id="firstname" value="<%=firstname%>" />
               	           <input name="lastname" type="hidden" id="lastname" value="<%=lastname%>" />
               	           <input name="stopped" type="hidden" id="stopped" value="<%=stopped%>" />
               	           <input name="remove" type="hidden" id="remove" value="<%=remove%>" />
               	           <input name="update" type="hidden" id="update" value="1" />
                                                  </p>
                	   </div>
                	 </form>
                	 <p align="center">&nbsp; </p>
            </div></td>
          </tr>
          <tr>
            <td scope="col"><div align="left"><p align="center">&nbsp;</p>
              </div></td>
          </tr>
        </table>
        <p align="center" class="contentHeader2">&nbsp;</p>
	</td>
    <td valign="top" bgcolor="#E7F1F7"><p align="center" class="contentHeader1">&nbsp;</p>
        <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><p>&nbsp;</p>
<p>MVH &Oslash;ko-net, <a href="mailto:kmh@eco-net.dk">kmh@eco-net.dk</a>, 62 24 43 24            </p>
            <p>&nbsp;</p></td>
          </tr>
        </table>
    </td>
  </tr>
</table>
</body>
</html>
