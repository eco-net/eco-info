<%
'**** Calculating sql-statement

DIM theFrom, theWhere, theOrder, searchDescr, eicount, eibackcount,einextcount
'theFrom="eiorg_maindata m"
theFrom="eiorg_maindata m LEFT JOIN eisys_postnr p ON m.zip_dk=p.postnr "
theWhere="(m.validated=1 or m.validated=0) AND (m.stopped=0 OR m.stopped=1)"

'theWhere="validated=1"
theOrder="ORDER BY p.postnr"
searchDescr=""

eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF request("sektion" & eicount)="dgb" THEN
	theWhere=theWhere & " AND lo.libid=" & request("recid" & eicount)
	theFROM= "(" & theFrom & " LEFT JOIN eilib_r_org lo ON m.id=lo.orgid)"
	searchDescr=searchDescr & " tilknyttet publikationen <b>" & request("recname" & eicount) & "</b>"
ELSEIF request("sektion" & eicount)="ok" THEN
	theWhere=theWhere & " AND lo.calid=" & request("recid" & eicount)
	theFROM= "(" & theFrom & " LEFT JOIN eical_r_org lo ON m.id=lo.orgid)"
	searchDescr=searchDescr & " tilknyttet arrangementet <b>" & request("recname" & eicount) & "</b>"
ELSEIF request("sektion" & eicount)="kurser" THEN
	theWhere=theWhere & " AND lo.courseid=" & request("recid" & eicount)
	theFROM= "(" & theFrom & " LEFT JOIN eicourse_r_org lo ON m.id=lo.orgid)"
	searchDescr=searchDescr & " tilknyttet kurset <b>" & request("recname" & eicount) & "</b>"
ELSEIF LEN(request("coll"))>0 THEN
	theWhere=theWhere & " AND c.siteid=1 AND c.collid=" & request("coll")
	theFROM= "(" & theFrom & " LEFT JOIN eiorg_r_coll c ON m.id=c.orgid)"
	theOrder="ORDER BY c.sortid"
	searchDescr="<b>" & request("collname") & "</b>"
	
ELSEIF LEN(request("new"))>0 THEN
	theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date),DatePart("d",date))
	theDate=DateAdd("d",(0-request("new")),theDate)
	theWhere=theWhere & " AND m.createdate>=" & FormatDateDB(theDate)
	searchDescr="oprettet indenfor de seneste <b>" & request("new") & " dage</b>"
ELSEIF LEN(request("id"))>0 THEN
	theWhere=theWhere & " AND m.id=" & request("id")
	searchDescr="<b>Specifik</b> organisation"
ELSE
	theCat=replace(request("dgscat"),"x","")
	theSbj=replace(request("dgssbj"),"x","")
	IF theCat="" THEN theCat=0
	IF theSbj="" THEN theSbj=0
	IF theCat>0 THEN
		IF InStr(request("dgscat"),"x")>0 THEN
			IF theCat=0.5 THEN
				theWhere=theWhere & " AND (cm.groupid=0 OR cm.groupid IS NULL)"
			ELSE
				theWhere=theWhere & " AND cm.groupid=" & theCat
			END IF
			theFrom= "((" & theFrom & " LEFT JOIN eiorg_r_cat oc ON m.id=oc.orgid) LEFT JOIN eiorg_cat_maindata cm ON oc.catid=cm.id)"
		ELSE
			theWhere=theWhere & " AND oc.catid=" & theCat
			theFROM= "(" & theFrom & " LEFT JOIN eiorg_r_cat oc ON m.id=oc.orgid)"
		END IF
		searchDescr=searchDescr & " i kategorien <b>" & request("dgscatname") & "</b>"
	END IF
	IF theSbj>0 THEN
		IF InStr(request("dgssbj"),"x")>0 THEN
			IF theSbj=0.5 THEN
				theWhere=theWhere & " AND (sm.groupid=0 OR sm.groupid IS NULL)"
			ELSE
				theWhere=theWhere & " AND sm.groupid=" & theSbj
			END IF
			theFrom= "((" & theFrom & " LEFT JOIN eiorg_r_sbj os ON m.id=os.orgid) LEFT JOIN eisbj_maindata sm ON os.sbjid=sm.id)"
		ELSE
			theWhere=theWhere & " AND os.sbjid=" & theSbj
			theFROM= "(" & theFrom & " LEFT JOIN eiorg_r_sbj os ON m.id=os.orgid)"
		END IF
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("dgssbjname") & "</b>"
	END IF
	IF request("dgskomm")>0 THEN
		theWhere=theWhere & " AND p.komnr=" & request("dgskomm")
		'theFROM= "(" & theFrom & " LEFT JOIN eisys_postnr p ON m.zip_dk=p.postnr)"
		searchDescr=searchDescr & " og hjemh&oslash;rende i kommunen <b>" & request("dgskommname") & "</b>"
	ELSEIF request("dgsamt")>0 THEN
		theWhere=theWhere & " AND p.amtnr=" & request("dgsamt")
		'theFROM= "(" & theFrom & " LEFT JOIN eisys_postnr p ON m.zip_dk=p.postnr)"
		searchDescr=searchDescr & " og hjemh&oslash;rende i amtet <b>" & request("dgsamtname") & "</b>"
	END IF
	IF LEN(request("dgstext"))>0 THEN
		theWhere=theWhere & " AND (m.name LIKE '%" & request("dgstext") &_
			"%' OR m.fullname LIKE '%" & request("dgstext") &_
			"%' OR m.shortdescription LIKE '%" & request("dgstext") &_
			"%' OR m.keywords LIKE '%" & request("dgstext") &_
			"%' OR m.description LIKE '%" & request("dgstext") & "%')"
		searchDescr=searchDescr & " og hvor navnet eller beskrivelsen indeholder teksten <b>" & request("dgsText") & "</b>"
	END IF
	IF LEN(request("key"))>0 THEN
		theWhere=theWhere & " AND (m.keywords LIKE '%" & request("key") & "%')"
		searchDescr=searchDescr & " og hvor navnet eller beskrivelsen indeholder teksten <b>" & request("key") & "</b>"
	END IF
END IF
IF searchDescr="" THEN
	searchDescr="Alle organisationer"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Organisationer " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE
		searchDescr="Organisationer " & searchDescr & "."
	END IF	
END IF


FUNCTION FormatDateDB(theDate)
	theDate=CDate(theDate)
	FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
		right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
END FUNCTION

'response.write theWhere & "<br>"
'response.write theFROM & "<br>"
'response.end
%>
