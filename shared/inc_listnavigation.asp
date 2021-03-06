<%
SetLocale("da")

'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variable

' set the record count
Function RegExpTest(patrn, strng)
   Dim regEx, Match, Matches   ' Create variable.
   Set regEx = New RegExp   ' Create a regular expression.
   regEx.Pattern = patrn   ' Set pattern.
   regEx.IgnoreCase = True   ' Set case insensitivity.
   regEx.Global = True   ' Set global applicability.
   Set Matches = regEx.Execute(strng)   ' Execute search.
   RegExpTest = Matches.Item(0).Value
End Function
set rsCount=server.createObject("ADODB.recordset")
rsCount.activeConnection=rsPageData.activeConnection
rsCount.source=replace(rsPageData.source,RegExpTest("SELECT[\d,\D]*(?=FROM){1}",rsPageData.source),"SELECT Count(*) as thenum ")
rsCount.source=replace(rsCount.source,RegExpTest("ORDER BY[\d,\D]*",rsCount.source),"")
rsCount.open()
rsPageData_total = rsCount.Fields.Item("thenum").value
rsCount.close()
rsCount=""

' set the number of rows displayed on this page
If (rsPageData_numRows < 0) Then
  rsPageData_numRows = rsPageData_total
Elseif (rsPageData_numRows = 0) Then
  rsPageData_numRows = 1
End If

' set the first and last displayed record
rsPageData_first = 1
rsPageData_last  = rsPageData_first + rsPageData_numRows - 1

' if we have the correct record count, check the other stats
If (rsPageData_total <> -1) Then
  If (rsPageData_first > rsPageData_total) Then rsPageData_first = rsPageData_total
  If (rsPageData_last > rsPageData_total) Then rsPageData_last = rsPageData_total
  If (rsPageData_numRows > rsPageData_total) Then rsPageData_numRows = rsPageData_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsPageData_total = -1) Then

  ' count the total records by iterating through the recordset
  rsPageData_total=0
  While (Not rsPageData.EOF)
    rsPageData_total = rsPageData_total + 1
    rsPageData.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsPageData.CursorType > 0) Then
    rsPageData.MoveFirst
  Else
    rsPageData.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsPageData_numRows < 0 Or rsPageData_numRows > rsPageData_total) Then
    rsPageData_numRows = rsPageData_total
  End If

  ' set the first and last displayed record
  rsPageData_first = 1
  rsPageData_last = rsPageData_first + rsPageData_numRows - 1
  If (rsPageData_first > rsPageData_total) Then rsPageData_first = rsPageData_total
  If (rsPageData_last > rsPageData_total) Then rsPageData_last = rsPageData_total

End If


' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsPageData
MM_rsCount   = rsPageData_total
MM_size      = rsPageData_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If


' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If


' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If

' *** Move To Record: update recordset stats

' set the first and last displayed record
rsPageData_first = MM_offset + 1
rsPageData_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsPageData_first > MM_rsCount) Then rsPageData_first = MM_rsCount
  If (rsPageData_last > MM_rsCount) Then rsPageData_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)


' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index=&offset="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function


' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "=##"


' Creating navigation-html
naviHtml="<a href=""index.asp"" title=""Tilbage til sektion-forsiden."">Til Forside</a> | "
IF MM_offset=0 THEN 
	naviHtml=naviHtml & "<img src=""/shared/graphics/layout/arrows_back.gif"" width=""10"" height=""7"" border=""0"">Forrige side"
ELSE
	naviHtml=naviHtml & "<a href=""" & replace(urlStr,"##",CStr(MM_offset - MM_size)) & """ title=""Til forrige side i oversigten""><img src=""/shared/graphics/layout/arrows_back.gif"" width=""10"" height=""7"" border=""0"">Forrige side</a>"
END IF
naviHtml=naviHtml & "&nbsp;|&nbsp;"
IF NOT MM_atTotal THEN
	naviHtml=naviHtml & "<a href=""" & replace(urlStr,"##",Cstr(MM_offset + MM_size)) & """ title=""Til n�ste side i oversigten"">N&aelig;ste side<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"" border=""0""></a><br>Side " 
ELSE
	naviHtml=naviHtml & "N&aelig;ste side<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"" border=""0""><br>Side "
END IF

pageCount=int(rsPageData_total/Repeat1__numRows) + 1
currentpagenum=int(MM_offset/Repeat1__numRows)+1

IF pageCount<=10 or currentpagenum<=5 THEN
	' there's only 10 or less pages in total, OR offset is within the first 5 pages.
	firstpageoffset=0
	firstpagenum=1
	lastpageoffset=rsPageData_total - ((rsPageData_total-1) mod Repeat1__numRows)
	IF pageCount>10 THEN lastpageoffset=MM_offset + ((10-currentpagenum)*Repeat1__numRows)
	lastpagenum=int(lastpageoffset/Repeat1__numRows) + 1
ELSE
	' There's more than 10 pages AND this is page 6 or more
	lastPageOffset=MM_offset + (5*Repeat1__numRows)
	IF lastpageoffset>=(rsPageData_total-1) THEN lastpageoffset=rsPageData_total -((rsPageData_total-1) mod Repeat1__numRows)
	firstpageoffset=lastpageoffset - 90
	firstpagenum=int(firstpageoffset/Repeat1__numRows)+1
	lastpagenum=int(lastpageoffset/Repeat1__numRows)+1
END IF

IF currentpagenum=1 THEN
	naviHtml=naviHtml & "<strong>1</strong>"
ELSE
	naviHtml=naviHtml & "<a href=""" & replace(urlStr,"##",Cstr(firstpageoffset)) & """ title=""Til denne side i oversigten"">" & firstpagenum & "</a>"
END IF

firstpagenum=firstpagenum+1
firstpageoffset=firstpageoffset+Repeat1__numRows

WHILE firstpagenum<=lastpagenum
	IF currentpagenum=firstpagenum THEN
		naviHtml=naviHtml & " | <strong>" & firstpagenum & "</strong>"
	ELSE
		naviHtml=naviHtml & " | <a href=""" & replace(urlStr,"##",Cstr(firstpageoffset)) & """ title=""Til denne side i oversigten"">" & firstpagenum & "</a>"	
	END IF
	firstpagenum=firstpagenum+1
	firstpageoffset=firstpageoffset+Repeat1__numRows
WEND

' Setting detail-link
detailLink="detail.asp?" & MM_keepBoth & MM_joinChar(MM_keepBoth) & "listoffset=" & MM_offset & "&offset=##"

'record info
recInfo="Nr. " & rsPageData_first & " - " & rsPageData_last & " vises."

%>
