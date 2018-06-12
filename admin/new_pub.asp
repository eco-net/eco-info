<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/admin/functions.asp"-->
<script src="/shared/multiselect.js"></script>
<script src="pub_lib.js"></script>
<script src="/shared/common.js"></script>
<script src="/shared/showeditor.js"></script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Ny publikation i &Oslash;ko-info</title>
<link href="/shared/styles.css" rel="stylesheet" type="text/css" />
<script type="text/JavaScript">
<!--

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
        <p align="center" class="contentHeader1">S&aring;dan</p>
        <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
          </tr>
          <tr>
            <td valign="top"><p>Opret en ny publikation ved at udfylde formularen. Brug de felter du finder 
n&oslash;dvendigt, men v&aelig;r opm&aelig;rksom p&aring;, at felter markeret 
med en * skal udfyldes.<br />
<br />
<span class="sidebarHeader">P&aring; nettet</span><br />
Udfyldes hvis publikationen er en web-side, eller hvis en n&aelig;rmere beskrivelse, 
bestilling o.l. kan foretages der.<br />
<br />
<span class="sidebarHeader">Om udgiveren</span><br />
Disse felter udfyldes kun, hvis du/I ikke selv har udgivet materialet.<br />
<span class="sidebarHeader"><br />
Kort og uddybende beskrivelse:</span><br />
Den korte beskrivelse vises i oversigten, mens b&aring;de kort og uddybende vises, 
n&aring;r nogen ser p&aring; dit kort i De Gr&oslash;nne Sider.</p>
                <p><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> <br />
                    <a href="#" onClick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">For 
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
      <form method="post" action="" onSubmit="donew(document.forms[0]);return document.validation">
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
              <tr>
                <td class="contentHeader1"> Ny Publikation </td>
              </tr>
              <tr>
                <td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1" /></td>
              </tr>
              <tr>
                <td valign="top"><table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr valign="top">
                        <td> Om Publikationen<br />
                            <span class="formLabeltext">*Titel:</span><br />
                            <input type="text" name="title" value="" size="32" class="formInputobjectLong" />
                            <br />
                            <span class="formLabeltext">Undertitel</span>: <br />
                            <input type="text" name="subtitle" value="" size="32" class="formInputobjectLong" />
                            <br />
                            <span class="formLabeltext">*Kategori:</span><br />
                            <select name="selCat" class="formInputobjectLong">
                              <script src="/log/insider/ei/menufiles/dgbcat_options.js"></script>
                            </select>
                            <br />
                          <span class="formLabeltext">*Forfatter:</span><br />
                          <input type="text" name="author" value="" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">ISBN:</span><br />
                          <input type="text" name="isbn" value="" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">Udgivelses&aring;r:</span><br />
                          <input type="text" name="year" value="" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">Sprog:</span><br />
                          <input type="text" name="language" value="" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">P&aring; nettet:</span><br />
                          <input type="text" name="webaddress" value="" size="32" class="formInputobjectLong" />
                        </td>
                        <td> Om udgiveren<br />
                            <span class="formLabeltext">Udgiver/Forlag:</span><br />
                            <textarea name="editor" cols="32" class="formTextobjectLow"></textarea>
                            <span class="formLabeltext"><br />
                              Email-adresse:</span><br />
                          <input type="text" name="editoremail" value="" size="32" class="formInputobjectLong" />
                          <span class="formLabeltext"><br />
                            P&aring; nettet:</span><br />
                          <input type="text" name="editorwww" value="" size="32" class="formInputobjectLong" /></td>
                      </tr>
                      <tr>
                        <td colspan="2">&nbsp;</td>
                      </tr>
                    </table>
                  <br />
                  Beskrivelser
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr valign="top">
                      <td width="50%"><span class="formLabeltext">*Kort beskrivelse:</span><br />
                      <textarea name="shortdescription" cols="50" rows="5" class="formTextobjectLow"></textarea>
                      <div align="center"></div></td>
                      <td><span class="formLabeltext">*Resum&eacute;:</span><br />
                      <textarea name="description" cols="50" rows="5" class="formTextobjectBig"></textarea>
                      <div align="center">
                        <input type="button" class="function" onClick="showeditor('description','');" value="Formater" name="button" />
                    </div></td>
                    </tr>
                  </table>
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
                        <textarea name="keywords" class="formTextobjectLow"></textarea></td>
                      <td><p>&nbsp;</p>
                      <p>&nbsp;</p></td>
                    </tr>
                    <tr valign="top">
                      <td><p><br />
                      </p>
                      </td>
                      <td><p>&nbsp;</p>
                      </td>
                    </tr>
                  </table>
                  <div align="center"><br />
                  <input name="reset" type="reset" class="formbutton" value="Ryd" />
                  <input name="submit" type="submit" class="formSubmitbutton" value="Opret" />
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
