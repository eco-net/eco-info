<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM eiorg_maindata  WHERE id=" & request("id")
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
%>
<%
Dim rsCategories__ColParam
rsCategories__ColParam = "0"
if (request("id") <> "") then rsCategories__ColParam = request("id")
%>
<%
set rsCategories = Server.CreateObject("ADODB.Recordset")
rsCategories.ActiveConnection = MM_ecoinfo_STRING
rsCategories.Source = "SELECT c.id,c.name FROM eiorg_r_cat o LEFT JOIN eiorg_cat_maindata c ON o.catid=c.id WHERE o.orgid=" + Replace(rsCategories__ColParam, "'", "''") + ""
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
rsSubjects.Source = "Select s.id,s.name FROM eiorg_r_sbj o LEFT JOIN eisbj_maindata s ON o.sbjid=s.id WHERE o.orgid=" + Replace(rsSubjects__ColParam, "'", "''") + ""
rsSubjects.CursorType = 0
rsSubjects.CursorLocation = 2
rsSubjects.LockType = 3
rsSubjects.Open()
rsSubjects_numRows = 0
%>
<%
Dim rsUserinf__MMColParam
rsUserinf__MMColParam = "0"
if (rsPageData.Fields.Item("id").value <> "") then rsUserinf__MMColParam = rsPageData.Fields.Item("id").value
%>
<%
set rsUserinf = Server.CreateObject("ADODB.Recordset")
rsUserinf.ActiveConnection = MM_ecoinfo_STRING
rsUserinf.Source = "SELECT *  FROM eisys_insiderusers  WHERE orgid = " + Replace(rsUserinf__MMColParam, "'", "''") + ""
rsUserinf.CursorType = 0
rsUserinf.CursorLocation = 2
rsUserinf.LockType = 3
rsUserinf.Open()
rsUserinf_numRows = 0
%>
<script src="/shared/multiselect.js"></script>
<script src="user_lib.js"></script>
<script src="/shared/common.js"></script>
<script src="/shared/showeditor.js"></script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Ret bruger i &Oslash;ko-info</title>
<link href="/shared/styles.css" rel="stylesheet" type="text/css" />
<script type="text/JavaScript">
<!--
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

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
function changepsw(){
un=document.form3.username.value;
psw=document.form3.password.value;
if (confirm("Vil du skifte til brugernavn: " + un + " og adgangskode: " + psw + "?"))
document.form3.submit();
}
function afmeld(){
if (confirm("Vil du afmelde?")){
document.form2.submit();
}
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
        <p align="center" class="contentHeader1">Ret bruger </p>
        <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
          </tr>
          <tr>
            <td valign="top"><p>&AElig;ndr de data du &oslash;nsker at &aelig;ndre, men v&aelig;r opm&aelig;rksom 
p&aring;, at felter markeret med en * skal v&aelig;re udfyldt.<br>
              <br>
              <span class="sidebarHeader">Navne</span><br>
              Hvis relevant indtastes navnet p&aring; organisationen / firmaet / projeket.<br>
              Vi skal altid have en kontakt-person.<br>
              <span class="sidebarHeader"><br>
                Kort og uddybende beskrivelse:</span><br>
              Den korte beskrivelse vises i oversigten, mens kun den uddybende vises, n&aring;r 
              nogen ser p&aring; dit kort i De Gr&oslash;nne Sider.</p>
                <p><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> <br>
                  Brug Formater-knappen<br>
                  MAC brugere:<br>
                  <a href="#" onClick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">For 
                    MAC brugere.</a><br>
                  <br>
                  <span class="sidebarHeader">Kategori og emneord:</span><br>
                  - er vigtige for at du bliver vist rette sted. Dine valg her er kun vejledende 
                  for vores redakt&oslash;rer der foretager den endelige kategorisering.<br>
                  <br>
                  <span class="sidebarHeader">Stikord</span><br>
                  P&aring; &Oslash;ko-info kan man lave fritekst-s&oslash;gninger. Hvis du &oslash;nsker 
                  at blive vist n&aring;r der s&oslash;ges p&aring; ord der ikke findes i dine beskrivelser, 
                  kan de skrives her.<br>
                  <br>
                  <span class="sidebarHeader">Brugernavn og adgangskode</span>:<br>
                  - de oplysninger du (og andre i din organisation) skal logge ind med for at bruge 
                  &Oslash;ko-info Insider.<br>
                  Den email-adresse der indtastes her, er den vi sender til, n&aring;r vi har brug 
                  for at kontakte dig/Jer.<br>
                  <br>
              </p></td>
          </tr>
          <tr>
            <td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
          </tr>
        </table>      
        <p align="center" class="contentHeader1">Ret Brugeroplysninger</p>
        <form id="form3" name="form3" method="post" action="newpsw.asp">
          <div align="center">
            <p>brugernavn
              <input name="username" type="text" id="username" />
              <br />
            adgangskode
            <input name="password" type="text" id="password" />
            <br />
            <input name="Ret" type="button" id="Ret" value="Ret" onClick="javascript:changepsw();"/>
            <input name="id" type="hidden" id="id" value="<%=rsPageData("id")%>" />
            <input name="name" type="hidden" id="name" value="<%=rsPageData("name")%>" />
            <input name="firstname" type="hidden" id="firstname" value="<%=rsPageData("firstname")%>" />
            <input name="lastname" type="hidden" id="lastname" value="<%=rsPageData("lastname")%>" />
            </p>
          </div>
        </form>
        <p align="center">&nbsp; </p>
        <p align="center"><span class="contentHeader1">Afmelding
        </span></p>
        <form id="form2" name="form2" method="post" action="afmeld.asp">
          <div align="center">
            <p align="left">
              <input name="stopped" type="checkbox" id="stopped" value="checkbox" />
            Org. er ophørt <br />
            <input name="remove" type="checkbox" id="remove" value="checkbox" />
            Registrering fjernes </p>
            <p>
              <input type="button" name="Submit2" value="Afmeld" onClick="afmeld()"/>
              <input name="id" type="hidden" id="id" value="<%=rsPageData("id")%>" />
              <input name="name" type="hidden" id="name" value="<%=rsPageData("name")%>" />
              <input name="firstname" type="hidden" id="firstname" value="<%=rsPageData("firstname")%>" />
              <input name="lastname" type="hidden" id="lastname" value="<%=rsPageData("lastname")%>" />
            </p>
          </div>
        </form>    </td>
    <td valign="top" bgcolor="#E7F1F7"><table width="100%" border="0" cellspacing="0" cellpadding="5" name="subIndexHeader">
      <form method="post" action="act_edit_user.asp" onSubmit="doedit(document.forms[0]);return document.validation" name="form1">
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
              <tr>
                <td class="contentHeader1">&nbsp;</td>
              </tr>
              <tr>
                <td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1" /></td>
              </tr>
              <tr>
                <td valign="top"><table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr valign="top">
                      <td> Organisation / Firma / Projekt<br />
                          <span class="formLabeltext">*Navn:</span><br />
                          <input type="text" name="name" value="<%=rsPageData("name")%>" size="32" class="formInputobjectLong" />
                          <br />
                          <br />                      </td>
                      <td> Kontaktperson <br />
                          <span class="formLabeltext">Titel:</span><br />
                          <input type="text" name="title" value="<%=rsPageData("title")%>" size="32" class="formInputobjectLong" />
                          <span class="formLabeltext"><br />
                            Fornavn:</span><br />
                        <input type="text" name="firstname" value="<%=rsPageData("firstname")%>" size="32" class="formInputobjectLong" />
                        <span class="formLabeltext"><br />
                          Efternavn:</span><br />
                        <input type="text" name="lastname" value="<%=rsPageData("lastname")%>" size="32" class="formInputobjectLong" />                      </td>
                    </tr>
                    <tr valign="top">
                      <td><br />
                        Adresse<br />
                        <span class="formLabeltext"> C/O (obs. C/O er automatisk foran):</span>&nbsp;&nbsp; <br />
                        <input type="text" name="adrco" value="<%=rsPageData("adrco")%>" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext"> *Gade:</span>&nbsp;&nbsp; <br />
                        <input type="text" name="street1" value="<%=rsPageData("street1")%>" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">Sted:&nbsp;&nbsp;</span> <br />
                        <input type="text" name="street2" value="<%=rsPageData("street2")%>" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">*Postnr:</span>&nbsp;&nbsp; (ikke by)<br />
                        <input type="text" name="zip" value="<%=rsPageData("zip")%>" size="32" class="formInputobjectLong" onChange="javascript:checkPostnr(this.value)" />
                        <input type="hidden" name="zip_dk" />                      </td>
                      <td><p><br />
                        Andre kontaktmuligheder <br />
                        <span class="formLabeltext"> Ved flere numre skriv da gerne lidt om nummeret i 
                          parentes: F.eks. xx xx xx xx (direkte/privat/mobil)<br />
                          Tlf 1:</span><br />
                        <input type="text" name="phone1" value="<%=rsPageData("phone1")%>" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">Tlf 2:</span><br />
                        <input type="text" name="phone2" value="<%=rsPageData("phone2")%>" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">Tlf 3:</span><br />
                        <input name="phone3" type="text" class="formInputobjectLong" id="phone3" value="<%=rsPageData("phone3")%>" size="32" />
                        <br />
                        <span class="formLabeltext">Fax:</span><br />
                        <input type="text" name="fax" value="<%=rsPageData("fax")%>" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">*Email addresse 1:</span><br />
                        <input type="text" name="emailaddress1" value="<%=rsPageData("emailaddress1")%>" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">Email addresse 2:</span><br />
                        <input type="text" name="emailaddress2" value="<%=rsPageData("emailaddress2")%>" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">www:</span> <br />
                        (www.hjemmeside.dk - skriv ikke http://)<br />
                        <input type="text" name="www" value="<%=rsPageData("www")%>" size="32" class="formInputobjectLong" />
                        <br />
                        www2: <br />
                        <input type="text" name="www2" value="<%=rsPageData("www2")%>" size="32" class="formInputobjectLong" />
                      </p></td>
                    </tr>
                  </table>
                  Beskrivelser
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr valign="top">
                      <td width="50%"><span class="formLabeltext">Kort beskrivelse:</span><br />
                        <textarea name="shortdescription" cols="50" rows="5" class="formTextobjectLow"><%=rsPageData("shortdescription")%></textarea>
                        <div align="center"></div></td>
                      <td><span class="formLabeltext">*Uddybende beskrivelse:</span><br />
                          <textarea name="description" cols="50" rows="5" class="formTextobjectBig"><%=rsPageData("description")%></textarea>
                          <div align="center">
                            <input type="button" class="function" onClick="showeditor('description','');" value="Formater" name="button" />
                        </div></td>
                    </tr>
                  </table>
                  *Kategorier
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr valign="top">
                      <td width="50%"><span class="formLabeltext">Alle kategorier:</span><br />
                          <select name="allCat" class="formTextobjectLow" size="5" ondblclick="addValue(this.form.allCat,this.form.selCat);">
                            <script src="/log/insider/ei/menufiles/dgscat_options.js"></script>
                          </select>
                          <br />
                        Dobbeltklik p&aring; en kategori for at v&aelig;lge den </td>
                      <td><span class="formLabeltext">Valgte kategorier:</span><br />
                        <select name="selCat" class="formTextobjectLow" size="5" ondblclick="removeValue(this.form.selCat);" multiple="multiple">
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
                        Dobbeltklik p&aring; en kategori for at fjerne den </td>
                    </tr>
                  </table>
                  <br />
                  *Emneord
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr valign="top">
                      <td width="50%"><span class="formLabeltext">Alle emneord:</span><br />
                      <select name="allSbj" class="formTextobjectLow" size="5"  ondblclick="addValue(this.form.allSbj,this.form.selSbj);">
                        <script src="/log/insider/ei/menufiles/sbj_options.js"></script>
                      </select>
                      <br />
                        Dobbeltklik p&aring; et emneord for at v&aelig;lge det </td>
                      <td><span class="formLabeltext">Valgte Emneord:</span><br />
                        <select name="selSbj" class="formTextobjectLow" size="5" ondblclick="removeValue(this.form.selSbj);" multiple="multiple">
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
                        <textarea name="keywords" class="formTextobjectLow"><%=rsPageData("keywords")%></textarea>                      </td>
                      <td><span class="formLabeltext"><br />
                      </span></td>
                    </tr>
                  </table>
                  <div align="center"><br />
                  <!--include virtual="/shared/inc_getparams.asp" -->
                  <%'=GetParams("&orgid=")%>
                  <input name="reset" type="reset" class="formbutton" value="Ryd" />
                  <input name="Submit" type="submit" class="formSubmitbutton" onClick="MM_validateForm('name','','R','street1','','R','zip','','R','emailaddress1','','R','description','','R');return document.MM_returnValue" value="Ret" />
                  </div></td>
              </tr>
          </table></td>
        </tr>
      </form>
    </table>    </td>
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
rsUserinf.Close()
%>
