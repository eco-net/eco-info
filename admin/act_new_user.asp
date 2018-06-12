<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp"-->
<!--#include virtual="/shared/common.asp"-->
<!--#include virtual="/admin/functions.asp"-->
<%
' NY BRUGER

name=request("name")
title=request("title")
firstname=request("firstname")
lastname=request("lastname")
adrco=request("adrco")
street1=request("street1")
street2=request("street2")
zip=request("zip")
phone1=request("phone1")
phone2=request("phone2")
phone3=request("phone3")
fax=request("fax")
emailaddress1=request("emailaddress1")
emailaddress2=request("emailaddress2")
www=request("www")
www2=request("www2")
shortdescription=request("shortdescription")
description=request("description")
keywords=request("keywords")
username=request("username")
password=request("password")
email=request("email")
'2. Inserting Categories.
thecat=" "
FOR each Item in request("selcat")
thecat=thecat & catsubname("cat",Item) & ", "
NEXT
'3. Inserting Subjects
thesub=" "
FOR each Item in request("selsbj")
thesub=thesub & catsubname("sbj",Item) & ", "
NEXT
style="<style type='text/css'><!--Body,Table {font-family: Verdana, Arial, Helvetica, sans-serif;	font-size: 11px;}--></style>"
userinput=style & "<table width=400 border=1 cellspacing=0 cellpadding=5><tr><td><b>Ny Bruger tilmeldt</b></td><td><b>Eco-info</b></td></tr><tr><td>Navn</td><td>" & name & "</td></tr><tr><td>Titel</td><td>" & title & "&nbsp;</td></tr><tr><td>Fornavn</td><td>" & firstname & "&nbsp;</td></tr><tr><td>Efternavn</td><td>" & lastname & "&nbsp;</td></tr><tr><td>C/O</td><td>" & adrco & "&nbsp;</td></tr><tr><td>Gade</td> <td>" & street1 & "&nbsp;</td></tr><tr><td>Sted</td><td>" & street2 & "&nbsp;</td></tr><tr><td>Postnr</td><td>" & zip & "&nbsp;</td><tr><tr><td>Tlf1</td><td>" & phone1 & "&nbsp;</td></tr><tr><td>Tlf2</td><td>" & phone2 & "&nbsp;</td></tr><tr><td>Tlf3</td><td>" & phone3 & "&nbsp;</td></tr><tr><td>Fax</td><td>" & fax & "&nbsp;</td></tr><tr><td>Email1</td><td>" & emailaddress1 & "&nbsp;</td></tr><tr><td>Email2</td><td>" & emailaddress2 & "&nbsp;</td></tr><tr><td>www</td><td>" & www & "&nbsp;</td></tr><tr><td>www2</td> <td>" & www2 & "&nbsp;</td></tr><tr><td>Kort beskrivelse </td><td>" & shortdescription & "&nbsp;</td></tr><tr><td>Uddybende beskrivelse </td><td>" & description & "&nbsp;</td></tr><tr><td>Kategorier</td><td>" & thecat & "&nbsp;</td></tr><tr><td>Emneord</td><td>" & thesub & "&nbsp;</td></tr><tr><td>Stikord</td><td>" & keywords & "&nbsp;</td></tr><tr><td>Brugernavn</td><td>" & username & "&nbsp;</td></tr><tr><td>&Oslash;nsket Adgangskode</td><td>" & password & "&nbsp;</td></tr><tr><td>Email-adresse</td><td>" & email & "&nbsp;</td></tr></table>"
theBody="Hej Kirsten<br>En bruger har oprettet sig på Øko-info " & Time() & "<br>Vil du oprette ham/hende i Insider/Filemaker tak" & userinput

SendCDOMail "kmh@eco-net.dk","web@eco-net.dk","NY BRUGER PÅ ØKO-INFO",theBody
SendCDOMail "eco-net@eco-net.dk","web@eco-net.dk","NY BRUGER PÅ ØKO-INFO",theBody
response.Redirect "thanks.asp?name=" & name & "&app=newuser"



%>