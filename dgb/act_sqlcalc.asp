<%
'**** Calculating sql-statement

DIM theFrom,theWhere,theOrder,searchDescr,eicount,eibackcount,einextcount
theFrom="eilib_maindata m"
theWhere="(validated=1 or validated=0) "
'theWhere="validated=1"
searchDescr=""
theOrder="ORDER BY m.title"
eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF request("sektion" & eicount)="dgs" THEN
	theWhere=theWhere & " AND lo.orgid=" & request("recid" & eicount)
	theFrom= "(" & theFrom & " LEFT JOIN eilib_r_org lo ON m.id=lo.libid)"
	theOrder="ORDER BY m.year DESC"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request("recname" & eicount) & "</b>"
ELSEIF LEN(request("coll"))>0 THEN
	theWhere=theWhere & " AND c.siteid=1 AND c.collid=" & request("coll")
	theFrom="(" & theFrom & " LEFT JOIN eilib_r_coll c ON m.id=c.libid)"
	theOrder="ORDER BY c.sortid"
	searchDescr="<b>" & request("collname") & "</b>"
ELSEIF LEN(request("new"))>0 THEN
	theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date),DatePart("d",date))
	theDate=DateAdd("d",(0-request("new")),theDate)
	theWhere=theWhere & " AND m.createdate>=" & FormatDateDB(theDate)
	searchDescr="oprettet indenfor de seneste <b>" & request("new") & " dage</b>"
ELSEIF LEN(request("id"))>0 THEN
	theWhere=theWhere & " AND m.id=" & request("id")
	searchDescr="<b>Specifik</b> publikation"
ELSEIF LEN(request("orgid"))>0 THEN
	theWhere=theWhere & " AND lo.orgid=" & request("orgid")
	theFrom= "(" & theFrom & " LEFT JOIN eilib_r_org lo ON m.id=lo.libid)"
	theOrder="ORDER BY m.year DESC"
	set rs=GetRecordSet("Select name from eiorg_maindata where id="&request("orgid"))
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & rs("name") & "</b>"
	set rs=nothing
ELSE
	theCat=replace(request("dgbcat"),"x","")
	theSbj=replace(request("dgbsbj"),"x","")
	IF theCat="" THEN theCat=0
	IF theSbj="" THEN theSbj=0
	IF theCat>0 THEN
		IF InStr(request("dgbcat"),"x")>0 THEN
			IF theCat=0.5 THEN
				theWhere=theWhere & " AND (cm.groupid=0 OR cm.groupid IS NULL)"
			ELSE
				theWhere=theWhere & " AND cm.groupid=" & theCat
			END IF
			theFrom= "((" & theFrom & " LEFT JOIN eilib_r_cat lc ON m.id=lc.libid) LEFT JOIN eilib_cat_maindata cm ON lc.catid=cm.id)"
		ELSE
			theWhere=theWhere & " AND lc.catid=" & theCat
			theFrom= "(" & theFrom & " LEFT JOIN eilib_r_cat lc ON m.id=lc.libid)"
		END IF
		searchDescr=searchDescr & " i kategorien <b>" & request("dgbcatname") & "</b>"
	END IF
	IF theSbj>0 THEN
		IF InStr(request("dgbsbj"),"x")>0 THEN
			IF theSbj=0.5 THEN
				theWhere=theWhere & " AND (sm.groupid=0 OR sm.groupid IS NULL)"
			ELSE
				theWhere=theWhere & " AND sm.groupid=" & theSbj
			END IF
			theFrom= "((" & theFrom & " LEFT JOIN eilib_r_sbj ls ON m.id=ls.libid) LEFT JOIN eisbj_maindata sm ON ls.sbjid=sm.id)"
		ELSE
			theWhere=theWhere & " AND ls.sbjid=" & theSbj
			theFrom= "(" & theFrom & " LEFT JOIN eilib_r_sbj ls ON m.id=ls.libid)"
		END IF
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("dgbsbjname") & "</b>"
	END IF
	IF LEN(request("dgbtext"))>0 THEN
		theWhere=theWhere & " AND (m.title LIKE '%" & request("dgbtext") &_
			"%' OR m.subtitle LIKE '%" & request("dgbtext") &_
			"%' OR m.author LIKE '%" & request("dgbtext") &_
			"%' OR m.shortdescription LIKE '%" & request("dgbtext") &_
			"%' OR m.description LIKE '%" & request("dgbtext") & "%')"
		searchDescr=searchDescr & " og hvor navnet eller beskrivelsen indeholder teksten <b>" & request("dgbText") & "</b>"
	END IF
	IF LEN(request("key"))>0 THEN
		theWhere=theWhere & " AND (m.keywords LIKE '%" & request("key") & "%')"
		searchDescr=searchDescr & " og hvor navnet eller beskrivelsen indeholder teksten <b>" & request("key") & "</b>"
	END IF
END IF
IF searchDescr="" THEN
	searchDescr="Alle publikationer"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Publikationer " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE	
		searchDescr="Publikationer " & searchDescr & "."
	END IF	
END IF

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

Function OrgFromLib(thelibid)
SET rsOFL = Server.CreateObject("ADODB.recordset")
	rsOFL.ActiveConnection=MM_ecoinfo_STRING
	rsOFL.Source="SELECT orgid FROM eilib_r_org WHERE libid=" & thelibid
	rsOFL.Open()
	orgid=rsOFL("orgid")
	rsOFL.Close
	OrgFromLib=orgid
end function

'response.write theWhere & "<br>"
'response.write theFROM & "<br>"
'response.end
%>
