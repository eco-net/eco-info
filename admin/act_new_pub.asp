<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp"-->
<!--#include virtual="/admin/functions.asp"-->
<!--#include virtual="/shared/common.asp"-->
<%
' NY PUBLIKATION

orgid=Session("orgid")
name=orgname(orgid)
title=request("title")
subtitle=request("subtitle")
shortdescription=request("shortdescription")
descr=request("description")
language=request("language")
author=request("author")
theyear=request("year")
isbn=request("isbn")
editor=request("editor")
editoremail=request("editoremail")
editorwww=request("editorwww")
webaddress=request("webaddress")
keywords=request("keywords")
selCat=catsubname("pubcat",request("selCat"))

'3. Inserting Subjects
thesub=" "
FOR each Item in request("selsbj")
thesub=thesub & catsubname("sbj",Item) & ", "
NEXT

style="<style type='text/css'><!--Body,Table {font-family: Verdana, Arial, Helvetica, sans-serif;	font-size: 11px;}--></style>"
userinput=style & "<table width=400 border=1 cellspacing=0 cellpadding=5><tr><td><b>Publikation oprettet</b></td><td><b>Eco-info</b></td></tr><tr><td>Titel</td><td>" & title & "</td></tr><tr><td>Undertitel</td><td>" & subtitle & "&nbsp;</td></tr><tr><td>Kategori</td><td>" & selCat & "&nbsp;</td></tr><tr><td>Forfatter</td><td>" & author & "&nbsp;</td></tr><tr><td>ISBN</td><td>" & isbn & "&nbsp;</td></tr><tr><td>Udgivelsesår</td> <td>" & theyear & "&nbsp;</td></tr><tr><td>Sprog</td><td>" & language & "&nbsp;</td></tr><tr><td>På nettet</td><td>" & webaddress & "&nbsp;</td><tr><tr><td>Udgiver/Forlag</td><td>" & editor & "&nbsp;</td></tr><tr><td>Email-adresse</td><td>" & editoremail & "&nbsp;</td></tr><tr><td>På nettet</td><td>" & editorwww & "&nbsp;</td></tr><tr><td>Kort beskrivelse </td><td>" & shortdescription & "&nbsp;</td></tr><tr><td>Resume</td><td>" & descr & "&nbsp;</td></tr><tr><td>Emneord</td><td>" & thesub & "&nbsp;</td></tr><tr><td>Stikord</td><td>" & keywords & "&nbsp;</td></tr></table>"
theBody="Hej Kirsten<br>" & name & " (orgid=" & orgid & ") har oprettet en publikation på Øko-info " & Time() & "<br>Vil du oprette den i Insider/Filemaker, tak" & userinput

SendCDOMail "kmh@eco-net.dk","web@eco-net.dk","NY PUBLIKATION PÅ ØKO-INFO",theBody
SendCDOMail "eco-net@eco-net.dk","web@eco-net.dk","NY PUBLIKATION PÅ ØKO-INFO",theBody
response.Redirect "thanks.asp?name=" & name & "&app=newpub&title=" & title


%>