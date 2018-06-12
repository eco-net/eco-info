<!--#include file="common.asp"-->
<%
theBody="En bruger er blevet afvist fra siden: " & request("p") & "<br>med brug af dette blacklistede ord: " & request("e")
SendCDOHtmlMail "hemure@gmail.com","eco-info@eco-net.dk","SQLERROR",theBody
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Fejl</title>
</head>

<body>
<h1 align="center">Fejl</h1>
<div style="background-color:#CC0000;color:#FFFFFF;font-size:18px;text-align:center">Fjern alle "<%=request("e")%>" fra teksten.</div>
<p align="center"><a href="javascript:history.go(-1)">Tilbage</a></p>
</body>
</html>
