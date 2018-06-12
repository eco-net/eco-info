<%
Dim navArray,p,nr
Function init()
Session("count")=0
Session("nav")=""
End Function

Function count(id)
Session("count")=Session("count")+1
if Session("nav")<> "" then 
Session("nav")= Session("nav") & Chr(44) & id 
else
Session("nav")= Session("nav") & id 
end if
End Function

Function getCount()
getCount=Session("count")
End Function

Function getNav()
getNav=Session("nav")
End Function

Function splitit()
navArray= Split(getnav(),Chr(44),-1,0)
End Function

Function thenext()
p=""
nr=request("id")
splitit()
i=0
For Each element in navArray
   if CStr(navArray(i))=CStr(nr) then
   		if i < Session("count") -1 then
  	 		p=navArray(i + 1)
		end if
  end if 
 i=i+1
Next 
thenext=p
End Function

Function theprev()
p=""
nr=request("id")
splitit()
i=0
For Each element in navArray
   if CStr(navArray(i))=CStr(nr) then
   		if i > 0 then
  	 		p=navArray(i - 1)
		end if
  end if 
 i=i+1
Next 
theprev=p
End Function
Function navHTML()
s1=""
s2=""
if theprev()<>"" then
s1="<a href='detail.asp?id=" & theprev() & "'>forrige</a> - "
end if
if thenext()<>"" then
s2="<a href='detail.asp?id=" & thenext() & "'>næste</a>"
end if
navHTML=s1+s2
End Function
%>

