<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/admin/functions.asp"-->
<!--#include virtual="/shared/common.asp"-->
<%
Dim rsPageData__ColParam
rsPageData__ColParam = "0"
if (request("id") <> "") then rsPageData__ColParam = request("id")
%>

<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM eical_maindata  WHERE id=" + Replace(rsPageData__ColParam, "'", "''") + ""
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<%
Dim rsCategories__ColParam
rsCategories__ColParam = "0"
if (request("id") <> "") then rsCategories__ColParam = request("id")
%>

<%
set rsCategories = Server.CreateObject("ADODB.Recordset")
rsCategories.ActiveConnection = MM_ecoinfo_STRING
rsCategories.Source = "SELECT c.id,c.name  FROM eical_cat_maindata c LEFT JOIN eical_r_cat o ON c.id=o.catid  WHERE o.calid=" + Replace(rsCategories__ColParam, "'", "''") + ""
rsCategories.CursorType = 0
rsCategories.CursorLocation = 2
rsCategories.LockType = 3
rsCategories.Open()
rsCategories_numRows = 0
%>
<%
Dim rsSubjects__ColParam
rsSubjects__ColParam = "0"
if (request("id") <> "") then rsSubjects__ColParam = request("id")
%>

<%
set rsSubjects = Server.CreateObject("ADODB.Recordset")
rsSubjects.ActiveConnection = MM_ecoinfo_STRING
rsSubjects.Source = "SELECT s.id,s.name  FROM eical_r_sbj o LEFT JOIN eisbj_maindata s ON o.sbjid=s.id  WHERE o.calid=" + Replace(rsSubjects__ColParam, "'", "''") + ""
rsSubjects.CursorType = 0
rsSubjects.CursorLocation = 2
rsSubjects.LockType = 3
rsSubjects.Open()
rsSubjects_numRows = 0
%>
<%
set rsAllCats = Server.CreateObject("ADODB.Recordset")
rsAllCats.ActiveConnection = MM_ecoinfo_STRING
rsAllCats.Source = "SELECT id,name  FROM eical_cat_maindata  WHERE siteid=1"
rsAllCats.CursorType = 0
rsAllCats.CursorLocation = 2
rsAllCats.LockType = 3
rsAllCats.Open()
rsAllCats_numRows = 0
%>
<script src="/shared/multiselect.js"></script>
<script src="arr_lib.js"></script>
<script src="/shared/common.js"></script>
<script src="/shared/showeditor.js"></script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Rediger arrangement i &Oslash;ko-info</title>
<link href="/shared/styles.css" rel="stylesheet" type="text/css" />
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body>
<table width="750" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#003300" bgcolor="#FFFFFF">
  <tr>
    <th colspan="2" bgcolor="#1979B5" scope="col"><div align="left"><img src="/shared/graphics/logo.gif" width="180" height="33" /></div></th>
  </tr>
  <tr>
    <td width="180" valign="top" bgcolor="#E7F1F7"><p align="center" class="contentHeader1">&nbsp;</p>
        <p align="center" class="contentHeader1">S&aring;dan</p>
        <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
          </tr>
          <tr>
            <td valign="top"><p>&AElig;ndr de data du &oslash;nsker at &aelig;ndre, men v&aelig;r opm&aelig;rksom 
p&aring;, at felter markeret med en * skal v&aelig;re udfyldt.<br />
<br />
<span class="sidebarHeader">Slutdato</span><br />
Udfyldes kun, hvis arrangementet str&aelig;kker sig over flere dage.</p>
                <p><span class="sidebarHeader">Arrang&oslash;rer</span> <br />
                  Der kan skrives flere arrang&oslash;rer og<br />
                  der kan linkes til flere organisationer i De Gr&oslash;nne Sider.</p>
                <p><span class="sidebarHeader">Sted og tid </span><br />
                  Skriv adresse og tid men ikke postnr/by.<br />
  <br />
  <span class="sidebarHeader">Postnr </span><br />
                  Skriv kun postnr, ikke by.<br />
  <br />
  <br />
  <span class="sidebarHeader">Beskrivelse:</span><br />
                  Beskrivelsen vises, n&aring;r nogen ser p&aring; dit kort i De Gr&oslash;nne Sider.<br />
                  Lange tekster kopieres ind,<br />
                  f.eks. fra et word dokument, da det er begr&aelig;nset hvor l&aelig;nge man kan 
                  v&aelig;re om at redigere.</p>
                <p><span class="sidebarHeader">Kort resume af beskrivelse:</span><br />
                  Vises kun i liste over arrangementer. Kan indeholde gentagelser fra beskrivelse.</p>
                <p><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> <br />
                    <a href="#" onClick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">For 
                      MAC brugere.</a><br />
  <br />
  <%IF LEN(request.cookies("eiuserid"))>0 THEN%>
  <span class="sidebarHeader">Kategori og emneord:</span><br />
                  - er vigtige for at du bliver vist rette sted. Dine valg her er kun vejledende 
                  for vores redakt&oslash;rer der foretager den endelige kategorisering.<br />
  <br />
  <%END IF%>
  <span class="sidebarHeader">Stikord</span><br />
                  P&aring; &Oslash;ko-info kan man lave fritekst-s&oslash;gninger. Hvis du &oslash;nsker 
                  at blive vist n&aring;r der s&oslash;ges p&aring; ord der ikke findes i dine beskrivelser, 
                  kan de skrives her.</p>
                <p><br>
                  <br>
              </p></td>
          </tr>
          <tr>
            <td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
          </tr>
        </table>      <p align="center" class="contentHeader1">&nbsp;</p></td>
    <td valign="top" bgcolor="#E7F1F7"><table width="100%" border="0" cellspacing="0" cellpadding="5" background="/dgs/graphics/sub_index_header_dgs_bkgrd.gif" name="subIndexHeader">
      <form method="post" action="" onSubmit="doedit(document.forms[0]);return document.validation">
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
              <tr>
                <td class="contentHeader1"> Ret arrangement</td>
              </tr>
              <tr>
                <td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1" /></td>
              </tr>
              <tr>
                <td valign="top"><table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr valign="top">
                      <td><p><span class="formLabeltext">*Titel:</span><br />
                        <input name="title" type="text" class="formInputobjectLong" value="<%=rsPageData("title")%>" size="32" />
                        <br />
                        <span class="formLabeltext">Undertitel</span>: <br />
                        <input type="text" name="subtitle" value="<%=rsPageData("subtitle")%>" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">*Starter den:</span><br />
                        <input type="text" name="startdate" value="<%=FormatDate(rsPageData("startdate"))%>" size="32" class="formInputobject" />
                        eks: 01-02-2002<br />
                        <span class="formLabeltext">Slutter den:</span><br />
                        <input type="text" name="enddate" value="<%=FormatDate(rsPageData("enddate"))%>" size="32" class="formInputobject" />
                        eks: 01-02-2002
                        
                        <br />
                        </p>
                        <strong>Arrangør: <%=orgname(Session("orgid"))%></strong>
                        <p><strong>Evt. medarrang&oslash;rer:</strong>                        <br />
                            <textarea name="organizers" cols="30" rows="3" id="organizers"><%=rsPageData("organizers")%></textarea>
                          </p></td>
                      <td><span class="formLabeltext">Starttidspunkt:</span><br />
                          <input name="starttime" type="text" class="formInputobjectLong" value="<%=rsPageData("starttime")%>" />
                          <br />
                          <span class="formLabeltext">Sluttidspunkt:</span><br />
                          <input name="endtime" type="text" class="formInputobjectLong" value="<%=rsPageData("endtime")%>" />
                          <br />
                          <span class="formLabeltext">Stedet for afholdelse:</span><br />
                          <input name="place" type="text" class="formInputobjectLong" value="<%=rsPageData("place")%>" />
                          <br />
                          <span class="formLabeltext">Adresse:</span><br />
                          <input name="address" type="text" class="formInputobjectLong" value="<%=rsPageData("address")%>" />
                          <br />
                          <span class="formLabeltext"> *Postnr:</span><br />
                          <input type="text" name="postnr" value="<%=rsPageData("postnr")%>" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext"> Telefon for n&aelig;rmere oplysninger:</span><br />
                          <input type="text" name="phone" value="<%=rsPageData("phone")%>" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">Email for n&aelig;rmere oplysninger:</span><br />
                          <input type="text" name="emailaddress" value="<%=rsPageData("emailaddress")%>" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">P&aring; nettet:</span> (www.hjemmeside.dk)<br />
                          <input type="text" name="www" value="<%=rsPageData("www")%>" size="32" class="formInputobjectLong" />
                      </td>
                    </tr>
                  </table>
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr valign="top">
                        <td width="50%"><span class="formLabeltext">*Beskrivelse:</span><br />
                            <textarea name="description" cols="30" rows="10"><%=rsPageData("description")%></textarea>
                            <div align="center">
                              <input type="button" class="function" onClick="showeditor('description','')"; value="Formater" name="button" />
                          </div></td>
                        <td><p align="left"><span class="formLabeltext">Kort resume af beskrivelse:</span><br />
                                <textarea name="shortdescription" cols="30" rows="5"><%=rsPageData("shortdescription")%></textarea>
                          </p>
                            <p>&nbsp; </p>
                            <div align="left"></div></td>
                      </tr>
                    </table>
                  <b><br />
                    *Kategorier</b>
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr valign="top">
                        <td width="50%"> Alle kategorier:<br />
                            <select name="allCat" class="formTextobjectLow" size="5"  ondblclick="addValue(this.form.allCat,this.form.selCat);">
                              <script src="/log/insider/ei/menufiles/okcat_options.js"></script>
                            </select>
                            <br />
                          Dobbeltklik p&aring; en kategori for at v&aelig;lge den</td>
                        <td> Valgte kategorier:<br />
                            <select name="selCat" size="5" multiple="MULTIPLE" class="formTextobjectLow" onDblClick="removeValue(this.form.selCat);">
							
<%
While (NOT rsCategories.EOF)
%>
<option value="<%=(rsCategories.Fields.Item("id").Value)%>"><%=(rsCategories.Fields.Item("name").Value)%></option>
<%
  rsCategories.MoveNext()
Wend
If (rsCategories.CursorType > 0) Then
  rsCategories.MoveFirst
Else
  rsCategories.Requery
End If
%>
</select>
                            <br />
                          Dobbeltklik p&aring; en kategori for at fjerne den</td>
                      </tr>
                    </table>
                  <br />
                    <b>*Emneord </b>
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr valign="top">
                        <td width="50%"> Alle emneord:<br />
                          <select name="allSbj" class="formTextobjectLow" size="5"  ondblclick="addValue(this.form.allSbj,this.form.selSbj);">
                            <script src="/log/insider/ei/menufiles/sbj_options.js"></script>
                          </select>
<br />
                          Dobbeltklik p&aring; et emneord for at v&aelig;lge det </td>
                        <td> Valgte Emneord:<br />
                            <select name="selSbj" size="5" multiple class="formTextobjectLow" onDblClick="removeValue(this.form.selSbj);">
<%
While (NOT rsSubjects.EOF)
%>
<option value="<%=(rsSubjects.Fields.Item("id").Value)%>"><%=(rsSubjects.Fields.Item("name").Value)%></option>
<%
  rsSubjects.MoveNext()
Wend
If (rsSubjects.CursorType > 0) Then
  rsSubjects.MoveFirst
Else
  rsSubjects.Requery
End If
%>
</select>

                            <br />
                          Dobbeltklik p&aring; et emneord for at fjerne det </td>
                      </tr>
                    </table>
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr valign="top">
                        <td width="50%"><span class="formLabeltext"><br />
                          Stikord:</span><br />
                          <textarea name="keywords" class="formTextobjectLow"><%=rsPageData("keywords")%></textarea>
                        </td>
                        <td><p>&nbsp;</p>
                        </td>
                      </tr>
                    </table>
                  <div align="center"><br />
                      <input name="reset" type="reset" class="formbutton" value="Ryd" />
                      <input name="Submit" type="submit" class="formSubmitbutton" onClick="MM_validateForm('title','','R','startdate','','R','postnr','','RisNum');return document.MM_returnValue" value="Ret" />
                  </div></td>
              </tr>
          </table></td>
        </tr>
      </form>
    </table></td>
  </tr>
</table>
</body>
</html>
<%
rsPageData.Close()
%>
<%
rsCategories.Close()
%>
<%
rsSubjects.Close()
%>
<%
rsAllCats.Close()
%>