<%
'**** Calculating sql-statement

DIM theFrom, theWhere, searchDescr, eicount, eibackcount,einextcount,theOrder
theFrom="eical_maindata m INNER JOIN eisys_postnr p ON m.postnr=p.postnr "
theWhere="(validated=1 or validated=0) "
theOrder="ORDER BY m.startdate"
searchDescr=""

eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF request("sektion" & eicount)="dgs" THEN
	theWhere=theWhere & " AND co.orgid=" & request("recid" & eicount)
	theFROM= "(" & theFrom & " LEFT JOIN eical_r_org co ON m.id=co.calid)"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request("recname" & eicount) & "</b>"
	theOrder="ORDER BY m.startdate DESC"
ELSEIF LEN(request("coll"))>0 THEN
	theWhere=theWhere & " AND c.siteid=1 AND c.collid=" & request("coll")
	theFROM= "(" & theFrom & " LEFT JOIN eical_r_coll c ON m.id=c.calid)"
	theOrder="ORDER BY c.sortid"
	searchDescr="<b>" & request("collname") & "</b>"
ELSEIF LEN(request("new"))>0 THEN
	theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date),DatePart("d",date))
	theDate=DateAdd("d",(0-request("new")),theDate)
	theWhere=theWhere & " AND m.createdate>=" & FormatDateDB(theDate)
	searchDescr="oprettet indenfor de seneste <b>" & request("new") & " dage</b>"
ELSEIF LEN(request("id"))>0 THEN
	theWhere=theWhere & " AND m.id=" & request("id")
	searchDescr="<b>Specifikt</b> arrangement"
ELSEIF LEN(request("orgid"))>0 THEN
	theWhere=theWhere & " AND co.orgid=" & request("orgid")
	theFROM= "(" & theFrom & " LEFT JOIN eical_r_org co ON m.id=co.calid)"
	set rs=GetRecordSet("Select name from eiorg_maindata where id="&request("orgid"))
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & rs("name") & "</b>"
	set rs=nothing
	theOrder="ORDER BY m.startdate DESC"

ELSE
	IF LEN(request("oktime"))>1 THEN
		DIM theDate
		IF CSTR(request("oktime"))="10" THEN
		'Kommende
			theWhere=theWhere & " AND (m.startdate>=" & FormatDateDB(date) &_
				" OR m.enddate>=" & FormatDateDB(date) & ")"
			searchDescr=searchDescr & "<b>Kommende</b> arrangementer"
		ELSEIF CSTR(request("oktime"))="20" THEN
		'Afholdte
			theWhere=theWhere & " AND m.startdate<=" & FormatDateDB(date) &_
				" AND m.enddate<=" & FormatDateDB(date)
			searchDescr=searchDescr & "<b>Afholdte</b> arrangementer"
			theOrder="ORDER BY m.startdate desc"
		ELSEIF CSTR(request("oktime"))="30" THEN
		'denne måned
			theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date)+1,1)
			theDate=DateAdd("d",-1,theDate)
			theDate=FormatDateDB(theDate)
			theWhere=theWhere & " AND (m.startdate BETWEEN " & FormatDateDB(Date) &_
				" AND " & theDate & " OR " &_
				"m.enddate BETWEEN " & FormatDateDB(Date) &_
				" AND " & theDate & ")"
			searchDescr=searchDescr & " i <b>resten af denne måned</b>"
		ELSEIF CSTR(request("oktime"))="40" THEN
		'næste måned
			DIM startDate
			startDate=DateSerial(DatePart("yyyy",date),DatePart("m",date)+1,1)
			theDate=DateAdd("m",1,startDate)
			startDate=FormatDateDB(startDate)
			theDate=DateAdd("d",-1,theDate)
			theDate=FormatDateDB(theDate)
			theWhere=theWhere & " AND (m.startdate BETWEEN " & startDate &_
				" AND " & theDate & " OR " &_
				"m.enddate BETWEEN " & startDate &_
				" AND " & theDate & ")"
			searchDescr=searchDescr & " i <b>næste måned</b>"
		ELSE
		'en måned
			theDate=DateAdd("m",1,CDate(request("oktime")))
			theDate=DateAdd("d",-1,theDate)
			theDate=FormatDateDB(theDate)
			theWhere=theWhere & " AND (m.startdate BETWEEN " & FormatDateDB(request("oktime")) &_
				" AND " & theDate & " OR " &_
				"m.enddate BETWEEN " & FormatDateDB(request("oktime")) &_
				" AND " & theDate & ")"
			searchDescr=searchDescr & " i perioden <b>" & request("oktimename") & "</b>"
		END IF
	ELSE
	'Kommende
		theWhere=theWhere & " AND (m.startdate>=" & FormatDateDB(date) &_
			" OR m.enddate>=" & FormatDateDB(date) & ")"
		searchDescr=searchDescr & "<b>Kommende</b> arrangementer"
	END IF

	theCat=replace(request("okcat"),"x","")
	theSbj=replace(request("oksbj"),"x","")
	IF theCat="" THEN theCat=0
	IF theSbj="" THEN theSbj=0

	IF theCat>0 THEN
		IF InStr(request("okcat"),"x")>0 THEN
			IF theCat=0.5 THEN
				theWhere=theWhere & " AND (cm.groupid=0 OR cm.groupid IS NULL)"
			ELSE
				theWhere=theWhere & " AND cm.groupid=" & theCat
			END IF
			theFrom= "((" & theFrom & " LEFT JOIN eical_r_cat cc ON m.id=cc.calid) LEFT JOIN eical_cat_maindata cm ON cc.catid=cm.id)"
		ELSE
			theWhere=theWhere & " AND cc.catid=" & theCat
			theFROM= "(" & theFrom & " LEFT JOIN eical_r_cat cc ON m.id=cc.calid)"
		END IF
		searchDescr=searchDescr & " i kategorien <b>" & request("okcatname") & "</b>"
	END IF
	IF theSbj>0 THEN
		IF InStr(request("oksbj"),"x")>0 THEN
			IF theSbj=0.5 THEN
				theWhere=theWhere & " AND (sm.groupid=0 OR sm.groupid IS NULL)"
			ELSE
				theWhere=theWhere & " AND sm.groupid=" & theSbj
			END IF
			theFrom= "((" & theFrom & " LEFT JOIN eical_r_sbj cs ON m.id=cs.calid) LEFT JOIN eisbj_maindata sm ON cs.sbjid=sm.id)"
		ELSE
			theWhere=theWhere & " AND cs.sbjid=" & theSbj
			theFROM= "(" & theFrom & " LEFT JOIN eical_r_sbj cs ON m.id=cs.calid)"
		END IF
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("oksbjname") & "</b>"
	END IF
	IF request("okkomm")>0 THEN
		theWhere=theWhere & " AND p.komnr=" & request("okkomm")
		searchDescr=searchDescr & " og afholdes i kommunen <b>" & request("okkommname") & "</b>"
	ELSEIF request("okamt")>0 THEN
		theWhere=theWhere & " AND p.amtnr=" & request("okamt")
		searchDescr=searchDescr & " og afholdes i amtet <b>" & request("okamtname") & "</b>"
	END IF
	IF LEN(request("oktext"))>0 THEN
		theWhere=theWhere & " AND (m.title LIKE '%" & request("oktext") &_
			"%' OR m.shortdescription LIKE '%" & request("oktext") &_
			"%' OR m.description LIKE '%" & request("oktext") & "%')"
		searchDescr=searchDescr & " og hvor titel eller beskrivelse indeholder teksten <b>" & request("okText") & "</b>"
	END IF
END IF
IF searchDescr="" THEN
	searchDescr="Alle arrangementer"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Arrangementer " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE
		searchDescr="Arrangementer " & searchDescr & "."
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

FUNCTION GetRecordSet(SQLstr)
	DIM theRS
	SET theRS = Server.CreateObject("ADODB.recordset")
	theRS.ActiveConnection=MM_ecoinfo_STRING
	theRS.Source=SQLstr
	theRS.Open()
	GetRecordSet=theRS
END FUNCTION

Function OrgFromCal(thecalid)
SET rsOFC = Server.CreateObject("ADODB.recordset")
	rsOFC.ActiveConnection=MM_ecoinfo_STRING
	rsOFC.Source="SELECT orgid FROM eical_r_org WHERE calid=" & thecalid
	rsOFC.Open()
	orgid=rsOFC("orgid")
	rsOFC.Close
	OrgFromCal=orgid
end function

%>
