<%

IF LEN(request("coll"))>0 THEN
'Back to index-page
	naviHtml="<a href=""index.asp"" title=""Tilbage til den side du kom fra"">Tilbage</a> | " & NaviHtml
ELSE
	naviHtml="<a href=""list.asp?" & backParams & """ title=""Tilbage til oversigten for din søgning"">Tilbage</a> | " & NaviHtml
END IF	

' Categories
rsData.source="SELECT c.name,c.id FROM eiopsl_cat_maindata c WHERE c.id=" & rsPageData.Fields.Item("cat_id").value
catList="<br><b>Findes i kategorierne:</b><br>"
rsData.open()
catList="<a href=""/opsl/list.asp?sektion=opsl&opslcat=" & rsData.Fields.Item("id").value & "&opslcatname=" & rsData.Fields.Item("name").value & """ title=""Finder opslag i denne kategori"">" & rsData.Fields.Item("name").value & "</a>"
rsData.close()
catList=catList & "<br>"
%>