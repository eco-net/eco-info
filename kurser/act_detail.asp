<%

forwardParams=MM_keepMove &_
	"&sektion" & einextcount & "=kurser" &_
	"&count=" & einextcount &_
	"&recid" & einextcount & "=" & rsPageData.Fields.Item("id").value &_
	"&recname" & einextcount & "=" & rsPageData.Fields.Item("title").value &_
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
	naviHtml="<a href=""/" & request("sektion" & request("count")) & "/detail.asp?" & backParams & """ title=""Tilbage til organisationen i De grønne sider"">Tilbage</a> | " & NaviHtml
END IF

' dgs - organisations
rsData.source="SELECT m.id,m.name,m.firstname,m.middlename,m.lastname,m.stopped FROM eiorg_maindata m LEFT JOIN eicourse_r_org o ON m.id=o.orgid WHERE o.courseid=" & rsPageData.Fields.Item("id").value
rsData.open()
IF NOT rsData.EOF THEN
	i=0
	dgsList="<br><table cellspacing=""0"" cellpadding=""0"" border=""0"">"
	WHILE NOT rsData.EOF
	if rsData.Fields.Item("stopped").value=1 then
		IF LEN(rsData.Fields.Item("name").value)>0 THEN
			dgsList=dgsList & "<tr valign='top'><td><a href='/dgs/stopped_detail.asp?id=" & rsData.Fields.Item("id").value & "'>" & rsData.Fields.Item("name").value & "</a></td></tr>"
		ELSE
			dgsList=dgsList & "<tr valign='top'><td><a href='/dgs/stopped_detail.asp?id=" & rsData.Fields.Item("id").value & "'>" & rsData.Fields.Item("firstname").value & "&nbsp;" & rsData.Fields.Item("middlename").value & "&nbsp;" & rsData.Fields.Item("lastname").value & "</a></td></tr>"
		END IF
	else
		IF LEN(rsData.Fields.Item("name").value)>0 THEN
			dgsList=dgsList & "<tr valign=""top""><td><a href=""/dgs/detail.asp?" & forwardParams & CStr(i) & """ title=""Viser denne organisation"">" & rsData.Fields.Item("name").value & "</a></td></tr>"
		ELSE
			dgsList=dgsList & "<tr valign=""top""><td><a href=""/dgs/detail.asp?" & forwardParams & CStr(i) & """ title=""Viser denne organisation"">" & rsData.Fields.Item("firstname").value & "&nbsp;" & rsData.Fields.Item("middlename").value & "&nbsp;" & rsData.Fields.Item("lastname").value & "</a></td></tr>"
		END IF
	end if
		rsData.moveNext
		i=i+1
	WEND
	dgsList=dgsList & "</table>"
END IF	
rsData.close()

' Categories
rsData.source="SELECT c.name,c.id FROM eicourse_cat_maindata c LEFT JOIN eicourse_r_cat l ON c.id=l.catid WHERE l.courseid=" & rsPageData.Fields.Item("id").value
catList="<br><b>Kategori:</b><br>"
rsData.open()
IF NOT rsData.EOF THEN
	catList=catList & "<a href=""/kurser/list.asp?cocat=" & rsData.Fields.Item("id").value & "&cocatname=" & rsData.Fields.Item("name").value & """ title=""Finder Kurser i denne kategori"">" & rsData.Fields.Item("name").value & "</a>"
	rsData.movenext
	WHILE NOT rsData.EOF
		catList=catList & " | <a href=""/okurser/list.asp?cocat=" & rsData.Fields.Item("id").value & "&cocatname=" & rsData.Fields.Item("name").value & """ title=""Finder Kurser i denne kategori"">" & rsData.Fields.Item("name").value & "</a>"
		rsData.movenext
	WEND
ELSE
	catList=catList & "Ingen"
END IF
rsData.close()
catList=catList & "<br>"
' Keywords
rsData.source="SELECT s.name,s.id FROM eisbj_maindata s LEFT JOIN eicourse_r_sbj l ON s.id=l.sbjid WHERE l.courseid=" & rsPageData.Fields.Item("id").value
sbjList="<br><b>Emneord:</b><br>"
rsData.open()
IF NOT rsData.EOF THEN
	sbjList=sbjList & "<a href=""/kurser/list.asp?cosbj=" & rsData.Fields.Item("id").value & "&cosbjname=" & rsData.Fields.Item("name").value & """ title=""Finder Kurser tilknyttet dette emne"">" & rsData.Fields.Item("name").value & "</a>"
	rsData.movenext
	WHILE NOT rsData.EOF
		sbjList=sbjList & " | <a href=""/kurser/list.asp?cosbj=" & rsData.Fields.Item("id").value & "&cosbjname=" & rsData.Fields.Item("name").value & """ title=""Finder Kurser tilknyttet dette emne"">" & rsData.Fields.Item("name").value & "</a>"
		rsData.movenext
	WEND
ELSE
	sbjList=sbjList & "Ingen"
END IF
rsData.close()
sbjList=sbjList & "<br>"

%>