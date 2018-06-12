<%
'Laver lister med relaterede data. Virker kun i sammenhæng med /shared/inc_detailnavigation.asp
forwardParams=MM_keepMove &_
	"&sektion" & einextcount & "=dgs" &_
	"&count=" & einextcount &_
	"&recid" & einextcount & "=" & rsPageData.Fields.Item("id").value &_
	"&recname" & einextcount & "=" & rsPageData.Fields.Item("name").value &_
	"&recoffset" & einextcount & "=" & MM_offset &_
	"&offset="
' Creating return-link
IF eibackcount="" THEN
' Back to list or indexpage
	IF LEN(request("coll"))>0 THEN
	'Back to index-page
		naviHtml="<a href=""index.asp"" title=""Tilbage til den side du kom fra"">Tilbage</a> | " & NaviHtml
	ELSE
		naviHtml="<a href=""list.asp?" & backParams & """ title=""Tilbage til oversigten for din søgning"">Tilbage</a> | " & NaviHtml
	END IF	
ELSE
	IF request("sektion" & eicount)="dgb" THEN
		toolText="Tilbage til publikationen i Det grønne bibliotek"
	ELSEIF request("sektion" & eicount)="ok" THEN
		toolText="Tilbage til arrangementet i Øko-kalenderen"
	END IF
	naviHtml="<a href=""/" & request("sektion" & eicount) & "/detail.asp?" & backParams & """ title=""" & toolText & """>Tilbage</a> | " & NaviHtml
END IF

'ok - events
rsData.source="SELECT m.id,m.title,m.startdate FROM eical_maindata m LEFT JOIN eical_r_org o ON m.id=o.calid WHERE o.orgid=" & rsPageData.Fields.Item("id").value & " ORDER BY m.startdate DESC"
rsData.open()
IF NOT rsData.EOF THEN
	i=0
	okList="<br><b>I &Oslash;ko-kalenderen:</b><br><table cellspacing=""0"" cellpadding=""0"" border=""0"">"
	WHILE NOT rsData.EOF
		okList=okList & "<tr valign=""top""><td nowrap>" & rsData.Fields.Item("startdate").value & "&nbsp;&nbsp;</td><td><a href=""/ok/detail.asp?" & forwardParams & Cstr(i) & """ title=""Viser dette arrangement"">" & rsData.Fields.Item("title").value & "</a></td></tr>"
		rsData.moveNext
		i=i+1
	WEND
	okList=okList & "</table>"
END IF
rsData.close()


' dgb - litterature
rsData.source="SELECT m.id,m.title,m.year FROM eilib_maindata m LEFT JOIN eilib_r_org o ON m.id=o.libid WHERE o.orgid=" & rsPageData.Fields.Item("id").value & "  ORDER BY year DESC"
rsData.open()
IF NOT rsData.EOF THEN
	i=0
		
	dgbList="<br><b>I Det Gr&oslashnne Bibliotek:</b><br><table cellspacing=""0"" cellpadding=""0"" border=""0"">"
	WHILE NOT rsData.EOF
		dgbList=dgbList & "<tr valign=""top""><td nowrap>" & rsData.Fields.Item("year").value & "&nbsp;&nbsp;</td><td><a href=""/dgb/detail.asp?" &_
			forwardParams & CStr(i) & """ title=""Viser denne publikation"">" & rsData.Fields.Item("title").value & "</a></td></tr>"
		rsData.moveNext
		i=i+1
	WEND
	dgbList=dgbList & "</table>"
END IF	
rsData.close()

' kurser
rsData.source="SELECT m.id,m.title,m.startdate FROM eicourse_maindata m LEFT JOIN eicourse_r_org o ON m.id=o.courseid WHERE o.orgid=" & rsPageData.Fields.Item("id").value & "  ORDER BY m.startdate DESC"
rsData.open()
IF NOT rsData.EOF THEN
	i=0
		
	dgbList="<br><b>Kurser / Uddannelser:</b><br><table cellspacing=""0"" cellpadding=""0"" border=""0"">"
	WHILE NOT rsData.EOF
		dgbList=dgbList & "<tr valign=""top""><td nowrap>" & rsData.Fields.Item("startdate").value & "&nbsp;&nbsp;</td><td><a href=""/kurser/detail.asp?" &_
			forwardParams & CStr(i) & """ title=""Viser dette kursus"">" & rsData.Fields.Item("title").value & "</a></td></tr>"
		rsData.moveNext
		i=i+1
	WEND
	dgbList=dgbList & "</table>"
END IF	
rsData.close()

' Categories
rsData.source="SELECT c.name,c.id FROM eiorg_cat_maindata c LEFT JOIN eiorg_r_cat o ON c.id=o.catid WHERE o.orgid=" & rsPageData.Fields.Item("id").value
catList="<br><b>Findes i kategorierne:</b><br>"
rsData.open()
IF NOT rsData.EOF THEN
	catList=catList & "<a href=""/dgs/list.asp?dgscat=" & rsData.Fields.Item("id").value & "&dgscatname=" & rsData.Fields.Item("name").value & """ title=""Finder organisationer i denne kategori"">" & rsData.Fields.Item("name").value & "</a>"
	rsData.movenext
	WHILE NOT rsData.EOF
		catList=catList & " | <a href=""/dgs/list.asp?dgscat=" & rsData.Fields.Item("id").value & "&dgscatname=" & rsData.Fields.Item("name").value & """ title=""Finder organisationer i denne kategori"">" & rsData.Fields.Item("name").value & "</a>"
		rsData.movenext
	WEND
ELSE
	catList=catList & "Ingen"
END IF
rsData.close()
catList=catList & "<br>"
' Keywords
rsData.source="SELECT s.name,s.id FROM eisbj_maindata s LEFT JOIN eiorg_r_sbj o ON s.id=o.sbjid WHERE o.orgid=" & rsPageData.Fields.Item("id").value
sbjList="<br><b>Emneord:</b><br>"
rsData.open()
IF NOT rsData.EOF THEN
	sbjList=sbjList & "<a href=""/dgs/list.asp?dgssbj=" & rsData.Fields.Item("id").value& "&dgssbjname=" & rsData.Fields.Item("name").value & """ title=""Finder organisationer tilknyttet dette emne"">" & rsData.Fields.Item("name").value & "</a>"
	rsData.movenext
	WHILE NOT rsData.EOF
		sbjList=sbjList & " | <a href=""/dgs/list.asp?dgssbj=" & rsData.Fields.Item("id").value& "&dgssbjname=" & rsData.Fields.Item("name").value & """ title=""Finder organisationer tilknyttet dette emne"">" & rsData.Fields.Item("name").value & "</a>"
		rsData.movenext
	WEND
ELSE
	sbjList=sbjList & "Ingen"
END IF
rsData.close()
sbjList=sbjList & "<br>"
%>	
