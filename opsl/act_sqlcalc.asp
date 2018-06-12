<%
'**** Calculating sql-statement

DIM theFrom, theWhere, searchDescr, eicount, eibackcount,einextcount,theOrder,thenew
theFrom="(eiopsl_maindata m LEFT JOIN eisys_postnr p ON m.postnr=p.postnr)"
theWhere="0=0"
theOrder="ORDER BY m.createdate DESC"
searchDescr=""
if request("new")<> "" then
thenew=0
else
thenew=365
end if
eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF LEN(request("coll"))>0 THEN
	theWhere=theWhere & " AND c.siteid=1 AND c.collid=" & request("coll")
	theFROM= "(" & theFrom & " LEFT JOIN eiopsl_r_coll c ON m.id=c.opslid)"
	theOrder="ORDER BY c.sortid"
	searchDescr="<b>" & request("collname") & "</b>"
ELSEIF LEN(request("new"))>0 THEN
	theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date),DatePart("d",date))
	theDate=DateAdd("d",(0-request("new")),theDate)
	theWhere=theWhere & " AND m.createdate>=" & FormatDateDB(theDate)
	searchDescr="oprettet indenfor de seneste <b>" & request("new") & " dage</b>"

ELSEIF LEN(request("id"))>0 THEN
	theWhere=theWhere & " AND m.id=" & request("id")
	searchDescr="<b>Specifikt</b> opslag"
ELSE
	theCat=replace(request("opslcat"),"x","")
	IF theCat="" THEN theCat=0

	IF theCat>0 THEN
		IF InStr(request("opslcat"),"x")>0 THEN
			IF theCat=0.5 THEN
				theWhere=theWhere & " AND (c.groupid=0 OR c.groupid IS NULL)"
			ELSE
				theWhere=theWhere & " AND c.groupid=" & theCat
			END IF
			theFrom= "(" & theFrom & " LEFT JOIN eiopsl_cat_maindata c ON m.cat_id=c.id)"
		ELSE
			theWhere=theWhere & " AND m.cat_id=" & theCat
		END IF
		searchDescr=searchDescr & " i kategorien <b>" & request("opslcatname") & "</b>"
	END IF
	IF request("opslkomm")>0 THEN
		theWhere=theWhere & " AND p.komnr=" & request("opslkomm")
		searchDescr=searchDescr & " og i kommunen <b>" & request("opslkommname") & "</b>"
	ELSEIF request("opslamt")>0 THEN
		theWhere=theWhere & " AND p.amtnr=" & request("opslamt")
		searchDescr=searchDescr & " og i amtet <b>" & request("opslamtname") & "</b>"
	END IF
	IF LEN(request("opsltext"))>0 THEN
		theWhere=theWhere & " AND (m.title LIKE '%" & request("opsltext") &_
			"%' OR m.kort_beskrivelse LIKE '%" & request("opsltext") &_
			"%' OR m.lang_beskrivelse LIKE '%" & request("opsltext") & "%')"
		searchDescr=searchDescr & " og hvor titel eller beskrivelse indeholder teksten <b>" & request("opslText") & "</b>"
	END IF
END IF
IF searchDescr="" THEN
	searchDescr="Alle opslag"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Opslag " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE
		searchDescr="Opslag " & searchDescr & "."
	END IF	
END IF
IF thenew >0 THEN
	theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date),DatePart("d",date))
	theDate=DateAdd("d",(0-thenew),theDate)
	theWhere=theWhere & " AND m.createdate>=" & FormatDateDB(theDate)
	searchDescr="Opslag oprettet indenfor de seneste <b>" & thenew & " dage</b>"
end if
'response.write theWhere & "<br>"
'response.write theFROM & "<br>"
'response.end

FUNCTION FormatDateDB(theDate)
	theDate=CDate(theDate)
	FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
		right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
END FUNCTION
%>
