<%

Function shortlink(link)
if InStr(link,"http://")>0 then
	if InStr(8,link,"/")>0 then
		a=InStr(8,link,"/")
	else
		a=Len(link)
	end if
else
	if InStr(link,"/")>0 then
		a=InStr(link,"/")
	else
		a=Len(link)
	end if
end if
shortlink=Left(link,a)

end Function

%>