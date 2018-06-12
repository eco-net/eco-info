<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp"-->
<!--#include virtual="/admin/functions.asp"-->
<!--#include virtual="/shared/common.asp"-->
<%
' NYT ARRANGEMENT

title=request("title")
subtitle=request("subtitle")
shortdescription=request("shortdescription")
description=request("description")
author=request("author")
theyear=request("year")
language=request("language")
webaddress=request("webaddress")
editor=request("editor")
editoremail=request("editoremail")
editorwww=request("editorwww")
orgid=Session("orgid")
name=orgname(orgid)
keywords=request("keywords")

'2. Inserting Categories.
thecat=" "
FOR each Item in request("selcat")
thecat=thecat & catsubname("artcat",Item) & ", "
NEXT
'3. Inserting Subjects
thesub=" "
FOR each Item in request("selsbj")
thesub=thesub & catsubname("sbj",Item) & ", "
NEXT

style="<style type='text/css'><!--Body,Table {font-family: Verdana, Arial, Helvetica, sans-serif;	font-size: 11px;}--></style>"
userinput=style & "<table width=400 border=1 cellspacing=0 cellpadding=5><tr><td><b>Artikel oprettet</b></td><td><b>Eco-info</b></td></tr><tr><td>Titel</td><td>" & title & "</td></tr><tr><td>Undertitel</td><td>" & subtitle & "&nbsp;</td></tr><tr><td>Kort beskrivelse </td><td>" & shortdescription & "&nbsp;</td></tr><tr><td>Uddybende beskrivelse </td><td>" & description & "&nbsp;</td></tr><tr><td>Forfatter</td><td>" & author & "&nbsp;</td></tr><tr><td>Udgivelsesdato</td><td>" & theyear & "&nbsp;</td></tr><tr><td>Sprog</td><td>" & language & "&nbsp;</td></tr><tr><td>Artikel-www</td><td>" & webaddress & "&nbsp;</td></tr><tr><td>Udgiver</td><td>" & editor & "&nbsp;</td></tr><tr><td>Udgiver-mail</td><td>" & editoremail & "&nbsp;</td></tr><tr><td>Udgiver-www</td><td>" & editorwww & "&nbsp;</td></tr><tr><td>Kategorier</td><td>" & thecat & "&nbsp;</td></tr><tr><td>Emneord</td><td>" & thesub & "&nbsp;</td></tr><tr><td>Stikord</td><td>" & keywords & "&nbsp;</td></tr></table>"
theBody="Hej Kirsten<br>" & name & " (orgid=" & orgid & ") har oprettet en artikel på Øko-info " & Time() & "<br>Vil du oprette den i Insider/Filemaker tak" & userinput

SendCDOMail "kmh@eco-net.dk","web@eco-net.dk","NY ARTIKEL PÅ ØKO-INFO",theBody
SendCDOMail "eco-net@eco-net.dk","web@eco-net.dk","NY ARTIKEL PÅ ØKO-INFO",theBody
response.Redirect "thanks.asp?name=" & name & "&app=newart&title=" & title


%>