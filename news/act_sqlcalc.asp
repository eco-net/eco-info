<%
Dim theWhere
theWhere= " s.econetsite<>1 "
if request("newscat")<>"" then
	theWhere=theWhere & " And ArtikelCategory_ID=" & CInt(request("newscat"))
	searchDescr="i kategorien <b>" & request("newscatname") & "</b>"
end if
if request("tema")<>"" then
	temaid=getTemaID()
	'response.write(temaid)
	'response.end
	theWhere=theWhere & " And Tema_ID=" & temaid 
	
end if
if request("searchtext")<>"" then
	s=toHTML(request("searchtext"))
	theWhere=theWhere & " And Header LIKE '%" & s & "%' or ShortDescription LIKE '%" & s & "%' or Content LIKE '%" & s & "%'" 
	searchDescr=searchDescr & " og hvor titel eller beskrivelse indeholder teksten <b>" & request("searchtext") & "</b>"
end if
if LEN(request("new"))>0 THEN
	theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date),DatePart("d",date))
	theDate=DateAdd("d",(0-request("new")),theDate)
	theWhere=theWhere & " AND create_date>=" & FormatDateDB(theDate)
	searchDescr="oprettet indenfor de seneste <b>" & request("new") & " dage</b>"
end if

IF searchDescr="" THEN
	searchDescr="Alle nyheder"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Nyheder " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE
		searchDescr="Nyheder " & searchDescr & "."
	END IF	
END IF

Function makeRed(str)
	search=request("searchtext")
	rep="<font color='red'>" & search & "</font>"
	makeRed=replace(str,search,rep)
end Function
FUNCTION FormatDateDB(theDate)
	theDate=CDate(theDate)
	FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
		right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
END FUNCTION
Function getTemaID()

	set con=Server.CreateObject("ADODB.Connection")
	set rs=Server.CreateObject("ADODB.Recordset")
	con.Open MM_ecoinfo_STRING
	
	theSQL="SELECT id FROM eitema_maindata WHERE chosen=1 "
	'rs.Open theSQL, con, 0,2,2
	rs.Open "eitema_maindata", con, 0,2,2
	while not rs.EOF
		if rs("chosen")=1 then
			theID=rs("id")
		end if
	rs.MoveNext
	wend
		
	rs.Close
	con.Close
	getTemaID=theID
End Function
%>