<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/admin/functions.asp"-->
<script src="/shared/multiselect.js"></script>
<script src="art_lib.js"></script>
<script src="/shared/common.js"></script>
<script src="/shared/showeditor.js"></script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Ny artikel i &Oslash;ko-info</title>
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
            <td valign="top"><p>Opret en ny Artikel ved at udfylde formularen. Brug de felter du finder n&oslash;dvendigt, 
men v&aelig;r opm&aelig;rksom p&aring;, at felter markeret med en * skal udfyldes.<br />
<br />
<span class="sidebarHeader">P&aring; nettet</span><br />
Udfyldes hvis artiklen er en web-side, eller hvis en n&aelig;rmere beskrivelse, 
bestilling o.l. kan foretages der.<br />
<br />
<span class="sidebarHeader">Om udgiveren</span><br />
Disse felter udfyldes kun, hvis du/I ikke selv har udgivet materialet.<br />
<span class="sidebarHeader"><br />
Kort beskrivelse:</span><br />
Den korte beskrivelse vises i kun i oversigten.</p>
                <p><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> <br />
                    <a href="#" onClick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">For 
                      MAC brugere.</a><br />
  <br />
  <span class="sidebarHeader">Kategori og emneord:</span><br />
                  - er vigtige for at artiklen bliver vist rette sted. Dine valg her er kun vejledende 
                  for vores redakt&oslash;rer der foretager den endelige kategorisering.<br />
  <br />
  <span class="sidebarHeader">Stikord</span><br />
                  P&aring; &Oslash;ko-info kan man lave fritekst-s&oslash;gninger. Hvis artiklen 
                  skal vises, n&aring;r der s&oslash;ges p&aring; ord der ikke findes i dine beskrivelser, 
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
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="5" background="/dgs/graphics/sub_index_header_dgs_bkgrd.gif" name="subIndexHeader">
            <tr>
              <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
                  <tr>
                    <td class="contentHeader1"> Ny Artikel</td>
                  </tr>
                  <tr>
                    <td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1" /></td>
                  </tr>
                  <tr>
                    <td valign="top"><table border="0" cellpadding="0" cellspacing="0" width="100%">
                          <tr>
                            <td class="sidebarHeaderBox">Artiklen</td>
                            <td class="sidebarHeaderBox">&nbsp;</td>
                          </tr>
                          <tr valign="top">
                            <td><p> <span class="formLabeltext">*Titel:</span><br />
                                    <input type="text" name="title" value="" size="32" class="formInputobjectLong" />
                                    <br />
                                    <span class="formLabeltext">Undertitel</span>: <br />
                                    <input type="text" name="subtitle" value="" size="32" class="formInputobjectLong" />
                                    <br />
                                    <br />
                                    <span class="formLabeltext">*Kort beskrivelse:</span><br />
                                    <textarea name="shortdescription" cols="25" rows="5"></textarea>
                                    <br />
                              </p>
                                <p align="center">&nbsp;</p>
                              <p align="center"><span class="listheader">Billeder:</span> <br />
                                Opret f&oslash;rst Artiklen<br />
                                Indsend derefter billeder</p></td>
                            <td><div align="center"><span class="formLabeltext">*Artikel-tekst:</span><br />
                                    <textarea name="description" cols="50" rows="5" class="formTextobjectBig"></textarea>
                                    <input type="button" class="function" onClick="showarteditor('description','');" value="Formater" name="button" />
                                    <span align="center"><br />
                                    <br />
                                    <span class="listheader">S&aring;dan g&oslash;r du</span></span><br />
                              Skriv artiklen uden formatering.<br />
                              Kopier den ind og brug Formater-knappen.<br />
                              Fjern evt. Word-kode inden formateringen.<br />
                              Siden kan h&oslash;jst v&aelig;re &aring;ben i 20 min.</div></td>
                          </tr>
                        </table>
                      <table border="0" cellpadding="0" cellspacing="0" width="100%">
                          <tr valign="top">
                            <td width="50%" class="sidebarHeaderBox">Forfatter / Udgiver</td>
                            <td class="sidebarHeaderBox">&nbsp;</td>
                          </tr>
                        </table>
                      <table border="0" cellpadding="0" cellspacing="0" width="100%">
                          <tr valign="top">
                            <td width="50%"><span class="formLabeltext">*Forfatter:</span><br />
                                <input type="text" name="author" size="32" class="formInputobjectLong" />
                                <br />
                                <span class="formLabeltext">Udgivelsesdato:</span><br />
                                <input type="text" name="year" size="32" class="formInputobjectLong" />
                                <br />
                                <span class="formLabeltext">Sprog:</span><br />
                                <input type="text" name="language" size="32" class="formInputobjectLong" />
                                <br />
                                <span class="formLabeltext">Artiklen findes p&aring; nettet:</span><br />
                              (www.hjemmeside.dk)<br />
                              <input type="text" name="webaddress" size="32" class="formInputobjectLong" />
                            </td>
                            <td><span class="formLabeltext">Udgiver/Forlag:</span><br />
                                <textarea name="editor" cols="32" class="formTextobjectLow"></textarea>
                                <span class="formLabeltext"><br />
                                  Email-adresse:</span><br />
                              <input type="text" name="editoremail" size="32" class="formInputobjectLong" />
                              <span class="formLabeltext"><br />
                                Udgiver p&aring; nettet:</span><br />
                              (www.hjemmeside.dk) <br />
                              <input type="text" name="editorwww" size="32" class="formInputobjectLong" />
                            </td>
                          </tr>
                        </table>
                      <table border="0" cellpadding="0" cellspacing="0" width="100%">
                          <tr valign="top">
                            <td width="50%" class="sidebarHeaderBox">Om Artiklen</td>
                            <td class="sidebarHeaderBox">&nbsp;</td>
                          </tr>
                        </table>
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td valign="top"><span class="formLabeltext">*Kategori:</span><br />
                                <br />
                                <select name="selCat" class="formInputobjectLong">
                                  <script src="/log/insider/ei/menufiles/artcat_options.js"></script>
                                </select>
                            </td>
                            <td><span class="formLabeltext">Stikord:</span><br />
                                <textarea name="keywords" class="formTextobjectLow"></textarea>
                            </td>
                          </tr>
                          <tr>
                            <td>*Emneord</td>
                            <td>&nbsp;</td>
                          </tr>
                          <tr>
                            <td><span class="formLabeltext">Alle emneord:</span><br />
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
                      <br />
                      <div align="center"><br />
                            <input name="reset" type="reset" class="formbutton" value="Ryd" />
                            <input name="submit" type="submit" class="formSubmitbutton" value="Opret" />
                        </div></td>
                  </tr>
              </table></td>
            </tr>
          </table></td>
        </tr>
      </form>
    </table></td>
  </tr>
</table>
</body>
</html>
