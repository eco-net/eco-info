<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp"-->
<!--#include virtual="/admin/functions.asp"-->
<!--#include virtual="/shared/common.asp"-->
<%
' NYT OPSLAG

titel=request("titel")

kort_beskrivelse=request("kort_beskrivelse")
lang_beskrivelse=request("lang_beskrivelse")
indsender=request("indsender")
email=request("email")
postnr=request("postnr")
telefon=request("telefon")
thecat=catsubname("opsl",request("cat_id"))

'response.write titel & "-" & kort_beskrivelse & "-" & lang_beskrivelse & "-" &  indsender & "-" & email & "-" & postnr & "-" & telefon & "-" & thecat
'response.end
style="<style type='text/css'><!--Body,Table {font-family: Verdana, Arial, Helvetica, sans-serif;	font-size: 11px;}--></style>"
userinput=style & "<table width=400 border=1 cellspacing=0 cellpadding=5><tr><td><b>Opslag oprettet</b></td><td><b>Eco-info</b></td></tr><tr><td>Titel</td><td>" & titel & "</td></tr><tr><td>Kort beskrivelse </td><td>" & kort_beskrivelse & "&nbsp;</td></tr><tr><td>Uddybende beskrivelse </td><td>" & lang_beskrivelse & "&nbsp;</td></tr><tr><td>Indsender</td><td>" & indsender & "&nbsp;</td></tr><tr><td>email</td><td>" & email & "&nbsp;</td></tr><tr><td>postnr</td><td>" & postnr & "&nbsp;</td></tr><tr><td>telefon</td><td>" & telefon & "&nbsp;</td></tr><tr><td>Kategori</td><td>" & thecat & "&nbsp;</td></tr></table>"
theBody="Hej Kirsten<br>" & indsender & " har oprettet et Opslag på Øko-info " & Time() & "<br>Vil du oprette det i Insider/Filemaker tak" & userinput

SendCDOMail "kmh@eco-net.dk","web@eco-net.dk","NYT OPSLAG PÅ ØKO-INFO",theBody
SendCDOMail "eco-net@eco-net.dk","web@eco-net.dk","NYT OPSLAG PÅ ØKO-INFO",theBody
response.Redirect "thanks.asp?name=" & indsender & "&app=newopsl&title=" & titel


%>