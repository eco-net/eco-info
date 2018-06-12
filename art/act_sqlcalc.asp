<%
'**** Calculating sql-statement

'DIM theFrom,theWhere
',theOrder,searchDescr,eicount,eibackcount,einextcount
theFrom="eiart_maindata m"
theWhere="0=0"
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

IF request("sektion" & eicount)="art" THEN
	theWhere=theWhere & " AND lo.orgid=" & request("recid" & eicount)
	theFrom= "(" & theFrom & " LEFT JOIN eiart_r_org lo ON m.id=lo.artid)"
	theOrder="ORDER BY m.authordate DESC"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request("recname" & eicount) & "</b>"
end if
if LEN(request("id"))>0 THEN
	theWhere=theWhere & " AND m.id=" & request("id")
	searchDescr="<b>Specifik</b> artikel"
end if
		if (request("artcat")>0) then
			theWhere=theWhere & " AND lc.catid=" & request("artcat")
			theFrom= "(" & theFrom & " LEFT JOIN eiart_r_cat lc ON m.id=lc.artid)"
		searchDescr=searchDescr & " i kategorien <b>" & request("artcatname") & "</b>"
		END IF
		if request("artsbj")>0 then
			theWhere=theWhere & " AND ls.sbjid=" & request("artsbj")
			theFrom= "(" & theFrom & " LEFT JOIN eiart_r_sbj ls ON m.id=ls.artid)"
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("artsbjname") & "</b>"
		END IF
	
		IF LEN(request("arttext"))>0 THEN
			theWhere=theWhere & " AND (m.title LIKE '%" & request("arttext") &_
				"%' OR m.subtitle LIKE '%" & request("arttext") &_
				"%' OR m.shortdescr LIKE '%" & request("arttext") &_
				"%' OR m.descr LIKE '%" & request("arttext") & "%')"
		searchDescr=searchDescr & " og hvor navnet eller beskrivelsen indeholder teksten <b>" & request("artText") & "</b>"
		END IF
		IF LEN(request("orgid"))>0 THEN
			theWhere=theWhere & " AND lo.orgid=" & request("orgid")
			theFrom= "(" & theFrom & " LEFT JOIN eiart_r_org lo ON m.id=lo.artid)"
		END IF

IF searchDescr="" THEN
	searchDescr="Alle artikler"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Artikler " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE	
		searchDescr="Artikler " & searchDescr & "."
	END IF	
END IF

FUNCTION FormatDateDB(theDate)
	theDate=CDate(theDate)
	FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
		right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
END FUNCTION

Function makeRed(str)
	search=request("arttext")
	'usearch=UCase(search)
	'lsearch=replace(search,Left(search,1),UCase(Left(search,1)))
	rep="<font color='red'>" & search & "</font>"
	makeRed=replace(str,search,rep)
end Function

Function OrgFromArt(theartid)
SET rsOFA = Server.CreateObject("ADODB.recordset")
	rsOFA.ActiveConnection=MM_ecoinfo_STRING
	rsOFA.Source="SELECT orgid FROM eiart_r_org WHERE artid=" & theartid
	rsOFA.Open()
	orgid=rsOFA("orgid")
	rsOFA.Close
	OrgFromArt=orgid
end function
%>
