<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp"-->
<!--#include virtual="/admin/functions.asp"-->
<!--#include virtual="/shared/common.asp"-->
<%
' NYT ARRANGEMENT

title=request("title")
subtitle=request("subtitle")
startdate=request("startdate")
enddate=request("enddate")
organizers=request("organizers")
starttime=request("starttime")
endtime=request("endtime")
place=request("place")
address=request("address")
postnr=request("postnr")
phone=request("phone")
emailaddress=request("emailaddress")
thewww=request("www")
orgid=Session("orgid")
name=orgname(orgid)
'name=""
shortdescription=request("shortdescription")
description=request("description")
keywords=request("keywords")

'2. Inserting Categories.
thecat=" "
FOR each Item in request("selcat")
thecat=thecat & catsubname("calcat",Item) & ", "
NEXT
'3. Inserting Subjects
thesub=" "
FOR each Item in request("selsbj")
thesub=thesub & catsubname("sbj",Item) & ", "
NEXT

style="<style type='text/css'><!--Body,Table {font-family: Verdana, Arial, Helvetica, sans-serif;	font-size: 11px;}--></style>"
userinput=style & "<table width=400 border=1 cellspacing=0 cellpadding=5><tr><td><b>Arrangement oprettet</b></td><td><b>Eco-info</b></td></tr><tr><td>Titel</td><td>" & title & "</td></tr><tr><td>Undertitel</td><td>" & subtitle & "&nbsp;</td></tr><tr><td>Startdato</td><td>" & FormatDate(startdate) & "&nbsp;</td></tr><tr><td>Slutdato</td><td>" & FormatDate(enddate) & "&nbsp;</td></tr><tr><td>Arrangører</td><td>" & organizers & "&nbsp;</td></tr><tr><td>Starttidspunkt</td> <td>" & starttime & "&nbsp;</td></tr><tr><td>Sluttidspunkt</td><td>" & endtime & "&nbsp;</td></tr><tr><td>Sted</td><td>" & place & "&nbsp;</td><tr><tr><td>Adresse</td><td>" & address & "&nbsp;</td></tr><tr><td>Postnr</td><td>" & postnr & "&nbsp;</td></tr><tr><td>Tlf</td><td>" & phone & "&nbsp;</td></tr><tr><td>Email1</td><td>" & emailaddress & "&nbsp;</td></tr><tr><td>www</td><td>" & thewww & "&nbsp;</td></tr><tr><td>Kort beskrivelse </td><td>" & shortdescription & "&nbsp;</td></tr><tr><td>Uddybende beskrivelse </td><td>" & description & "&nbsp;</td></tr><tr><td>Kategorier</td><td>" & thecat & "&nbsp;</td></tr><tr><td>Emneord</td><td>" & thesub & "&nbsp;</td></tr><tr><td>Stikord</td><td>" & keywords & "&nbsp;</td></tr></table>"
theBody="Hej Kirsten<br>" & name & " (orgid=" & orgid & ") har oprettet et arrangement på Øko-info " & Time() & "<br>Vil du oprette ham/hende i Insider/Filemaker tak" & userinput

SendCDOMail "kmh@eco-net.dk","web@eco-net.dk","NYT ARRANGEMENT PÅ ØKO-INFO",theBody
SendCDOMail "eco-net@eco-net.dk","web@eco-net.dk","NYT ARRANGEMENT PÅ ØKO-INFO",theBody
response.Redirect "thanks.asp?name=" & name & "&app=newarr"


%>