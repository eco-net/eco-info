<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%
if Session("stemmedeltager")<>Session.SessionID then
Session("stemmedeltager")=Session.SessionID
else
response.redirect("only_one.asp")
end if
%>
<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_ecoinfo_STRING
  MM_editTable = "eiafstem_ning"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "textfield|value|textfield2|value|afsid|value"
  MM_columnsStr = "forslag|',none,''|descr|',none,''|emne_id|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
if request("mail")<>"din@mailadresse" and request("mail")<>"" then
theSQL="INSERT INTO eiafstem_konkurrence (emne_id,mail) VALUES (" &_
	request("afsid") & "," &_
	"'" & replace(request("mail"),"'","''") & "'" &_
	")"
	execCommand theSQL
end if
End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
set rsAfstemning = Server.CreateObject("ADODB.Recordset")
rsAfstemning.ActiveConnection = MM_ecoinfo_STRING
rsAfstemning.Source = "SELECT n.id, n.forslag, n.stemmer, n.descr, e.id as eid, e.name, e.descr as d,e.forslag as ef  FROM eiafstem_ning n LEFT JOIN eiafstem_emne e ON n.emne_id=e.id  WHERE e.chosen=1  ORDER BY n.stemmer desc"
rsAfstemning.CursorType = 0
rsAfstemning.CursorLocation = 2
rsAfstemning.LockType = 3
rsAfstemning.Open()
rsAfstemning_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsAfstemning_numRows = rsAfstemning_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td> 
<div align="center" class="contentHeader2"><%=(rsAfstemning.Fields.Item("name").Value)%></div>
<br>
<div align="center"> 
<p><%=(rsAfstemning.Fields.Item("d").Value)%></p>
<p>&nbsp; </p>
</div>
</td>
</tr>
<tr> 
<td> <% if CStr(rsAfstemning.Fields.Item("ef").Value)="0" then %> 
<form name="form1" method="POST" action="<%=MM_editAction%>">
<div align="center"> 
<p><span class="formLabeltext">Skriv dit forslag til afstemningen:</span><br>
<input type="text" name="textfield" class="formInputobject" maxlength="30">
max 30 tegn<br>
<span class="formLabeltext">Beskrivelse:</span><br>
<input type="text" name="textfield2" class="formInputobjectLong" maxlength="60">
max 60 tegn</p>
<p>Deltag i konkurrencen om Bogen<br>
&Oslash;kologi er p&aring; alles l&aelig;ber:<br>
<input type="text" name="mail" class="formInputobject" value="din@mailadresse">
<br>
<input type="submit" name="Submit" value="Indsend" class="formSubmitbutton">
</p>
<p>&nbsp; </p>
<p> 
<input type="hidden" name="MM_insert" value="true">
<input type="hidden" name="afsid" value="<%=(rsAfstemning.Fields.Item("eid").Value)%>">
</p>
</div>
</form>
<%end if %>
</td>
</tr>
</table>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td> 
<form name="form2" action="stem2.asp" method="post">
<div align="center"> 
<table width="100%" border="0" cellspacing="0" cellpadding="3">
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsAfstemning.EOF)) 
%>
<tr> 
<td><%=(rsAfstemning.Fields.Item("forslag").Value)%></td>
<td><i><%=(rsAfstemning.Fields.Item("descr").Value)%></i></td>
<td><%=(rsAfstemning.Fields.Item("stemmer").Value)%></td>
<td><img src="/shared/graphics/layout/Oe.gif" width="25" height="27"></td>
<td> 
<input type="checkbox" name="checkbox" value="<%=(rsAfstemning.Fields.Item("id").Value)%>">
</td>
</tr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsAfstemning.MoveNext()
Wend
%>
</table>
<p>Deltag i konkurrencen om bogen:<br>
<i><b><a href="http://www.eco-info.dk/dgb/detail.asp?id=5907" target="_blank">&Oslash;kologi 
er p&aring; alles l&aelig;ber</a></b></i><br>
<input type="text" name="mail2" class="formInputobject" value="din@mailadresse">
<br>
<input type="submit" name="Submit2" value="Stem" class="formSubmitbutton">
<input type="hidden" name="stem" value="1">
<%rsAfstemning.MoveFirst %>
<input type="hidden" name="afsid2" value="<%=(rsAfstemning.Fields.Item("eid").Value)%>">
</p>
</div>
</form>
</td>
</tr>
</table>
<p>&nbsp;</p>
<span class="contentHeader2"></span> 
</body>
</html>
<%
rsAfstemning.Close()
%>
