<%
'**** Calculating sql-statement

DIM theFrom, theWhere, searchDescr, eicount, eibackcount,einextcount,theOrder,validtext
validtext=""
'theFrom="eicourse_maindata m"
theFrom="eicourse_maindata m INNER JOIN eisys_postnr p ON p.postnr = m.postnr "
'theFrom="(eisys_postnr p INNER JOIN eicourse_maindata m ON p.postnr = m.postnr) INNER JOIN (eicourse_cat_maindata c INNER JOIN eicourse_r_cat r ON c.id = r.catid) ON m.id = r.courseid;"
theWhere="0=0"
theOrder="ORDER BY m.startdate"

searchDescr=""

eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF request("sektion" & eicount)="dgs" THEN
	theWhere=theWhere & " AND co.orgid=" & request("recid" & eicount)
	theFROM= "(" & theFrom & " LEFT JOIN eicourse_r_org co ON m.id=co.courseid)"
	theOrder="ORDER BY m.startdate DESC"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request("recname" & eicount) & "</b>"
ELSEIF LEN(request("coll"))>0 THEN
	theWhere=theWhere & " AND c.siteid=1 AND c.collid=" & request("coll")
	theFROM= "(" & theFrom & " LEFT JOIN eicourse_r_coll c ON m.id=c.courseid)"
	theOrder="ORDER BY c.sortid"
	searchDescr="<b>" & request("collname") & "</b>"
ELSEIF LEN(request("new"))>0 THEN
	theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date),DatePart("d",date))
	theDate=DateAdd("d",(0-request("new")),theDate)
	theWhere=theWhere & " AND m.createdate>=" & FormatDateDB(theDate)
	searchDescr="oprettet indenfor de seneste <b>" & request("new") & " dage</b>"
ELSEIF LEN(request("id"))>0 THEN
	theWhere=theWhere & " AND m.id=" & request("id")
	searchDescr="<b>Specifikt</b> kursus"
ELSE
	
	IF LEN(request("cotime"))>1 THEN
		DIM theDate
		IF CSTR(request("cotime"))="10" THEN
		'Kommende
			theWhere=theWhere & " AND (m.startdate>=" & FormatDateDB(date) &_
				" OR m.enddate>=" & FormatDateDB(date) & ")"
			searchDescr=searchDescr & "<b>Igangværende og kommende</b> kurser"
		ELSEIF CSTR(request("cotime"))="20" THEN
		'Afholdte
			theWhere=theWhere & " AND m.startdate<=" & FormatDateDB(date) &_
				" AND m.enddate<=" & FormatDateDB(date)
			searchDescr=searchDescr & "<b>Afholdte</b> kurser"
		ELSEIF CSTR(request("cotime"))="30" THEN
		'denne måned
			theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date)+1,1)
			theDate=DateAdd("d",-1,theDate)
			'theDate=theDate
			theWhere=theWhere & " AND (m.startdate BETWEEN " & Date &_
				" AND " & theDate & " OR " &_
				"m.enddate BETWEEN " & Date &_
				" AND " & theDate & ")"
			searchDescr=searchDescr & " i <b>resten af denne måned</b>"
		ELSEIF CSTR(request("cotime"))="40" THEN
		'næste måned
			DIM startDate
			startDate=DateSerial(DatePart("yyyy",date),DatePart("m",date)+1,1)
			theDate=DateAdd("m",1,startDate)
			startDate=startDate
			theDate=DateAdd("d",-1,theDate)
			theDate=theDate
			theWhere=theWhere & " AND (m.startdate BETWEEN " & startDate &_
				" AND " & theDate & " OR " &_
				"m.enddate BETWEEN " & startDate &_
				" AND " & theDate & ")"
			searchDescr=searchDescr & " i <b>næste måned</b>"
		ELSE
		'en måned
			theDate=DateAdd("m",1,CDate(request("cotime")))
			theDate=DateAdd("d",-1,theDate)
			'theDate=FormatDateDB(theDate)
			theWhere=theWhere & " AND (m.startdate BETWEEN " & request("cotime") &_
				" AND " & theDate & " OR " &_
				"m.enddate BETWEEN " & request("cotime") &_
				" AND " & theDate & ")"
			searchDescr=searchDescr & " i perioden <b>" & request("cotimename") & "</b>"
		END IF
	ELSE
	'Kommende
		theWhere=theWhere & " AND (m.startdate>=" & FormatDateDB(date) &_
			" OR m.enddate>=" & FormatDateDB(date) & ")"
		searchDescr=searchDescr & "<b>Igangværende og kommende</b> kurser"
	END IF

	IF request("cocat")>0 THEN
		theWhere=theWhere & " AND cc.catid=" & request("cocat")
		theFrom= "(" & theFrom & " LEFT JOIN eicourse_r_cat cc ON m.id=cc.courseid)"
		searchDescr=searchDescr & " i kategorien <b>" & request("cocatname") & "</b>"
	END IF
	IF request("cosbj")>0 THEN
		theWhere=theWhere & " AND cs.sbjid=" & request("cosbj")
		theFrom= "(" & theFrom & " LEFT JOIN eicourse_r_sbj cs ON m.id=cs.courseid)"
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("cosbjname") & "</b>"
	END IF
	IF request("cokomm")>0 THEN
		theWhere=theWhere & " AND p.komnr=" & request("cokomm")
		searchDescr=searchDescr & " og afholdes i kommunen <b>" & request("cokommname") & "</b>"
	ELSEIF request("coamt")>0 THEN
		theWhere=theWhere & " AND p.amtnr=" & request("coamt")
		searchDescr=searchDescr & " og afholdes i amtet <b>" & request("coamtname") & "</b>"
	END IF
	IF LEN(request("cotext"))>0 THEN
		theWhere=theWhere & " AND (m.title LIKE '%" & request("cotext") &_
			"%' OR m.shortdescription LIKE '%" & request("cotext") &_
			"%' OR m.description LIKE '%" & request("cotext") & "%')"
		searchDescr=searchDescr & " og hvor titel eller beskrivelse indeholder teksten <b>" & request("coText") & "</b>"
	END IF
END IF

IF searchDescr="" THEN
	searchDescr="Alle kurser"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Kurser " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE
		searchDescr="Kurser " & searchDescr & "."
	END IF	
END IF
'response.write theWhere & "<br>"
'response.write theFROM & "<br>"
'response.end
FUNCTION FormatDateDB(theDate)
	theDate=CDate(theDate)
	FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
		right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
END FUNCTION

%>
