<%
SetLocale(1030)
Function daysleft(theday)
Dim slutdag,dage
slutdag=CDate(theday)
daysleft=DateDiff("d", Date, slutdag)
end function
%>