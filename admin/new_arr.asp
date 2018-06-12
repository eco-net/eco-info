<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/admin/functions.asp"-->
<script src="/shared/multiselect.js"></script>
<script src="arr_lib.js"></script>
<script src="/shared/common.js"></script>
<script src="/shared/showeditor.js"></script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Nyt arrangement i &Oslash;ko-info</title>
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
            <td valign="top"><p>Opret et nyt arrangement ved at udfylde formularen. Brug de felter du finder 
n&oslash;dvendige, men v&aelig;r opm&aelig;rksom p&aring;, at felter markeret 
med en * skal udfyldes.<br />
<br />
<span class="sidebarHeader">Slutdato</span><br />
Udfyldes kun, hvis arrangementet str&aelig;kker sig over flere dage.<br />
<br />
<span class="sidebarHeader">Arrang&oslash;rer</span> <br />
Der kan skrives flere arrang&oslash;rer og<br />
der kan linkes til flere organisationer i De Gr&oslash;nne Sider.<br />
<br />
<span class="sidebarHeader">Sted og tid </span><br />
Skriv adresse og tid men ikke postnr/by.<br />
<br />
<span class="sidebarHeader">Postnr </span><br />
Skriv kun postnr, ikke by.<br />
<br />
<span class="sidebarHeader">Beskrivelse:</span><br />
Beskrivelsen vises, n&aring;r nogen ser p&aring; dit kort i De Gr&oslash;nne Sider.<br />
Lange tekster kopieres ind,<br />
f.eks. fra et word dokument, da det er begr&aelig;nset hvor l&aelig;nge man kan 
v&aelig;re om at redigere.</p>
                <p><span class="sidebarHeader">Kort resume af beskrivelse:</span><br />
                  Vises kun i liste over arrangementer. Kan indeholde gentagelser fra beskrivelse.</p>
                <p><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> <br />
                    <a href="#" onclick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">For 
                      MAC brugere.</a><br />
  <br />
  <span class="sidebarHeader">Kategori og emneord:</span><br />
                  - er vigtige for at du bliver vist rette sted. Dine valg her er kun vejledende 
                  for vores redakt&oslash;rer der foretager den endelige kategorisering.<br />
  <br />
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
      <form method="post" action="" onsubmit="donew(document.forms[0]);return document.validation">
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
              <tr>
                <td class="contentHeader1"> Nyt arrangement</td>
              </tr>
              <tr>
                <td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1" /></td>
              </tr>
              <tr>
                <td valign="top"><table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr valign="top">
                      <td><p><span class="formLabeltext">*Titel:</span><br />
                        <input type="text" name="title" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">Undertitel</span>: <br />
                        <input type="text" name="subtitle" value="" size="32" class="formInputobjectLong" />
                        <br />
                        <span class="formLabeltext">*Starter den:</span><br />
                        <input type="text" name="startdate" value="" size="32" class="formInputobject" />
                        eks: 01-02-2002<br />
                        <span class="formLabeltext">Slutter den:</span><br />
                        <input type="text" name="enddate" value="" size="32" class="formInputobject" />
                        eks: 01-02-2002
                        
                        <br />
                        </p>
                        <strong>Arrangør: <%=orgname(Session("orgid"))%></strong>
                        <p><strong>Evt. medarrang&oslash;rer:</strong>                        <br />
                            <textarea name="organizers" cols="30" rows="3" id="organizers"></textarea>
                                              </p></td>
                      <td><span class="formLabeltext">Starttidspunkt:</span><br />
                          <input type="text" name="starttime" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">Sluttidspunkt:</span><br />
                          <input type="text" name="endtime" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">Stedet for afholdelse:</span><br />
                          <input type="text" name="place" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">Adresse:</span><br />
                          <input type="text" name="address" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext"> *Postnr:</span><br />
                          <input type="text" name="postnr" value="" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext"> Telefon for n&aelig;rmere oplysninger:</span><br />
                          <input type="text" name="phone" value="" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">Email for n&aelig;rmere oplysninger:</span><br />
                          <input type="text" name="emailaddress" value="" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">P&aring; nettet:</span> (www.hjemmeside.dk)<br />
                          <input type="text" name="www" value="" size="32" class="formInputobjectLong" />
                      </td>
                    </tr>
                  </table>
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr valign="top">
                        <td width="50%"><span class="formLabeltext">*Beskrivelse:</span><br />
                            <textarea name="description" cols="30" rows="10"></textarea>
                            <div align="center">
                              <input type="button" class="function" onclick="showeditor('description','')"; value="Formater" name="button" />
                          </div></td>
                        <td><p align="left"><span class="formLabeltext">Kort resume af beskrivelse:</span><br />
                                <textarea name="shortdescription" cols="30" rows="5"></textarea>
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
                            <select name="selCat" class="formTextobjectLow" size="5" ondblclick="removeValue(this.form.selCat);" multiple="multiple">
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
                        <td><p>&nbsp;</p>
                        </td>
                      </tr>
                    </table>
                  <div align="center"><br />
                      <input name="reset" type="reset" class="formbutton" value="Ryd" />
                      <input name="Submit" type="submit" class="formSubmitbutton" onclick="MM_validateForm('title','','R','startdate','','R','postnr','','RisNum');return document.MM_returnValue" value="Opret" />
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
