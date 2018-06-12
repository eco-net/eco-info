<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<script src="/shared/multiselect.js"></script>
<script src="user_lib.js"></script>
<script src="/shared/common.js"></script>
<script src="/shared/showeditor.js"></script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Ny bruger i &Oslash;ko-info</title>
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
        <p align="center" class="contentHeader1">Ny bruger </p>
        <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
          </tr>
          <tr>
            <td valign="top"><p>Tilmeld dig De gr&oslash;nne Sider ved at udfylde formularen. Brug de felter 
              du finder n&oslash;dvendigt, men v&aelig;r opm&aelig;rksom p&aring;, at felter 
              markeret med en * skal udfyldes.<br>
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
        </table>      <p align="center" class="contentHeader1">&nbsp;</p></td>
    <td valign="top" bgcolor="#E7F1F7"><table width="100%" border="0" cellspacing="0" cellpadding="5" name="subIndexHeader">
      <form method="post" action="act_new_user.asp" onSubmit="donew(document.forms[0]);return document.validation" name="form1">
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
                          <input type="text" name="name" value="" size="32" class="formInputobjectLong" />
                          <br />
                          <br />
                      </td>
                      <td> Kontaktperson <br />
                          <span class="formLabeltext">Titel:</span><br />
                          <input type="text" name="title" value="" size="32" class="formInputobjectLong" />
                          <span class="formLabeltext"><br />
                            Fornavn:</span><br />
                        <input type="text" name="firstname" value="" size="32" class="formInputobjectLong" />
                        <span class="formLabeltext"><br />
                          Efternavn:</span><br />
                        <input type="text" name="lastname" value="" size="32" class="formInputobjectLong" />
                      </td>
                    </tr>
                    <tr valign="top">
                      <td><br />
                        Adresse<br />
                        <span class="formLabeltext"> C/O (obs. C/O er automatisk foran):</span>&nbsp;&nbsp; <br />
                        <input type="text" name="adrco" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext"> *Gade:</span>&nbsp;&nbsp; <br />
                        <input type="text" name="street1" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">Sted:&nbsp;&nbsp;</span> <br />
                        <input type="text" name="street2" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">*Postnr:</span>&nbsp;&nbsp; (ikke by)<br />
                        <input type="text" name="zip" value="" size="32" class="formInputobjectLong" onChange="javascript:checkPostnr(this.value)" />
                        <input type="hidden" name="zip_dk" />
                      </td>
                      <td><p><br />
                        Andre kontaktmuligheder <br />
                        <span class="formLabeltext"> Ved flere numre skriv da gerne lidt om nummeret i 
                          parentes: F.eks. xx xx xx xx (direkte/privat/mobil)<br />
                          Tlf 1:</span><br />
                        <input type="text" name="phone1" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">Tlf 2:</span><br />
                        <input type="text" name="phone2" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">Tlf 3:</span><br />
                        <input name="phone3" type="text" class="formInputobjectLong" id="phone3" value="" size="32" />
                        <br />
                        <span class="formLabeltext">Fax:</span><br />
                        <input type="text" name="fax" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">*Email addresse 1:</span><br />
                        <input type="text" name="emailaddress1" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">Email addresse 2:</span><br />
                        <input type="text" name="emailaddress2" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">www:</span> <br />
                        (www.hjemmeside.dk - skriv ikke http://)<br />
                        <input type="text" name="www" value="" size="32" class="formInputobjectLong" />
                        <br />
                        www2: <br />
                        <input type="text" name="www2" value="" size="32" class="formInputobjectLong" />
                      </p></td>
                    </tr>
                  </table>
                  Beskrivelser
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr valign="top">
                      <td width="50%"><span class="formLabeltext">Kort beskrivelse:</span><br />
                          <textarea name="shortdescription" cols="50" rows="5" class="formTextobjectLow"></textarea>
                          <div align="center"></div></td>
                      <td><span class="formLabeltext">*Uddybende beskrivelse:</span><br />
                          <textarea name="description" cols="50" rows="5" class="formTextobjectBig"></textarea>
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
                      </select>
                      <br />
                        Dobbeltklik p&aring; et emneord for at fjerne det </td>
                    </tr>
                  </table>
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr valign="top">
                      <td width="50%"><span class="formLabeltext"><br />
                        Stikord:</span><br />
                        <textarea name="keywords" class="formTextobjectLow"></textarea>
                      </td>
                      <td><span class="formLabeltext"><br />
                        *Brugernavn:</span><br />
                        <input type="text" name="username" value="" size="32" class="formInputobjectLong" />
                        <span class="formLabeltext"><br />
                          *&Oslash;nsket Adgangskode:</span><br />
                        <input type="text" name="password" value="" size="32" class="formInputobjectLong" />
                        <span class="formLabeltext"><br />
                          *Email-adresse:</span><br />
                        <input type="text" name="email" value="" size="32" class="formInputobjectLong" />
                      </td>
                    </tr>
                  </table>
                  <div align="center"><br />
                  <!--include virtual="/shared/inc_getparams.asp" -->
                  <%'=GetParams("&orgid=")%>
                  <input name="reset" type="reset" class="formbutton" value="Ryd" />
                  <input name="Submit" type="submit" class="formSubmitbutton" onClick="MM_validateForm('name','','R','street1','','R','zip','','R','emailaddress1','','R','username','','R','password','','R','email','','RisEmail','description','','R');return document.MM_returnValue" value="Opret" />
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
