<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/admin/functions.asp"-->

<%
Dim rsOrg__ColParam
rsOrg__ColParam = "0"
if (request("id") <> "") then rsOrg__ColParam = request("id")
%>

<%
set rsOrg = Server.CreateObject("ADODB.Recordset")
rsOrg.ActiveConnection = MM_ecoinfo_STRING
rsOrg.Source = "SELECT o.id,o.name, o.firstname,o.lastname  FROM eiorg_maindata o LEFT JOIN eiart_r_org l ON o.id=l.orgid  WHERE l.artid=" + Replace(rsOrg__ColParam, "'", "''") + ""
rsOrg.CursorType = 0
rsOrg.CursorLocation = 2
rsOrg.LockType = 3
rsOrg.Open()
rsOrg_numRows = 0
%>

<%
Dim rsPageData__ColParam
rsPageData__ColParam = "0"
if (request("id") <> "") then rsPageData__ColParam = request("id")
%>
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM eiart_maindata  WHERE id=" + Replace(rsPageData__ColParam, "'", "''") + ""
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
rsCategories.Source = "SELECT c.id,c.name  FROM eiart_cat_maindata c LEFT JOIN eiart_r_cat o ON c.id=o.catid  WHERE o.artid=" + Replace(rsCategories__ColParam, "'", "''") + ""
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
rsSubjects.Source = "SELECT s.id,s.name  FROM eiart_r_sbj o LEFT JOIN eisbj_maindata s ON o.sbjid=s.id  WHERE o.artid=" + Replace(rsSubjects__ColParam, "'", "''") + ""
rsSubjects.CursorType = 0
rsSubjects.CursorLocation = 2
rsSubjects.LockType = 3
rsSubjects.Open()
rsSubjects_numRows = 0
%>
<%
set rsAllCats = Server.CreateObject("ADODB.Recordset")
rsAllCats.ActiveConnection = MM_ecoinfo_STRING
rsAllCats.Source = "SELECT id,name  FROM eiart_cat_maindata"
rsAllCats.CursorType = 0
rsAllCats.CursorLocation = 2
rsAllCats.LockType = 3
rsAllCats.Open()
rsAllCats_numRows = 0
%>


<script src="/shared/multiselect.js"></script>
<script src="art_lib.js"></script>
<script src="/shared/common.js"></script>
<script src="/shared/showeditor.js"></script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Rediger artikel i &Oslash;ko-info</title>
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
            <td valign="top"><p>&AElig;ndr de data du &oslash;nsker at &aelig;ndre, men v&aelig;r opm&aelig;rksom 
p&aring;, at felter markeret med en * skal v&aelig;re udfyldt.<br />
<br />
<span class="sidebarHeader">P&aring; nettet</span><br />
Udfyldes hvis publikationen er en web-side, eller hvis en n&aelig;rmere beskrivelse, 
bestilling o.l. kan foretages der.</p>
                <p><span class="sidebarHeader">Om udgiveren</span><br />
                  Disse felter udfyldes kun, hvis du/I ikke selv har udgivet materialet.</p>
                <p><span class="sidebarHeader">Kort og uddybende beskrivelse</span><br />
                  Den korte beskrivelse vises i oversigten, mens b&aring;de kort og uddybende vises, 
                  n&aring;r nogen ser p&aring; dit kort i De Gr&oslash;nne Sider.</p>
                <p><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> <br />
                    <a href="#" onclick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">Se 
                      her hvordan.</a></p>
                <p><span class="sidebarHeader">Kategori og emneord</span><br />
                  - er vigtige for at du bliver vist rette sted. Dine valg her er kun vejledende 
                  for vores redakt&oslash;rer der foretager den endelige kategorisering.</p>
                <p><span class="sidebarHeader">Stikord</span><br />
                  P&aring; &Oslash;ko-info kan man lave fritekst-s&oslash;gninger. Hvis du &oslash;nsker 
                  at blive vist n&aring;r der s&oslash;ges p&aring; ord der ikke findes i dine beskrivelser, 
                  kan de skrives her.<br />
                </p>
                <p><br>
                  </p></td>
          </tr>
          <tr>
            <td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
          </tr>
        </table>      <p align="center" class="contentHeader1">&nbsp;</p></td>
    <td valign="top" bgcolor="#E7F1F7"><table width="100%" border="0" cellspacing="0" cellpadding="5" background="/dgs/graphics/sub_index_header_dgs_bkgrd.gif" name="subIndexHeader">
      <form method="post" action="" onSubmit="doedit(document.forms[0]);return document.validation">
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="5" background="/dgs/graphics/sub_index_header_dgs_bkgrd.gif" name="subIndexHeader">
            <tr>
              <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
                  <tr>
                    <td class="contentHeader1"> Ret Artikel</td>
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
                                    <input type="text" name="title" value="<%=rsPageData("title")%>" size="32" class="formInputobjectLong" />
                                    <br />
                                    <span class="formLabeltext">Undertitel</span>: <br />
                                    <input type="text" name="subtitle" value="<%=rsPageData("subtitle")%>" size="32" class="formInputobjectLong" />
                                    <br />
                                    <br />
                                    <span class="formLabeltext">*Kort beskrivelse:</span><br />
                           <textarea name="shortdescription" cols="25" rows="5">
						   <%=rsPageData("shortdescr")%> 
						   </textarea>
                                    <br />
                              </p>
                                <p align="center">&nbsp;</p>
                              <p align="center"><span class="listheader">Billeder:</span> <br />
                                Opret f&oslash;rst Artiklen<br />
                                Indsend derefter billeder</p></td>
                            <td><div align="center"><span class="formLabeltext">*Artikel-tekst:</span><br />
                                    <textarea name="description" cols="50" rows="5" class="formTextobjectBig">
									<%=rsPageData("descr")%>
									</textarea>
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
                                <input name="author" type="text" class="formInputobjectLong" value="<%=rsPageData("author")%>" size="32" />
                                <br />
                                <span class="formLabeltext">Udgivelses&aring;r:</span><br />
                                <input name="year" type="text" class="formInputobjectLong" value="<%=rsPageData("authordate")%>" size="32" />
                                <br />
                                <span class="formLabeltext">Sprog:</span><br />
                                <input name="language" type="text" class="formInputobjectLong" value="<%=rsPageData("lang")%>" size="32" />
                                <br />
                                <span class="formLabeltext">Artiklen findes p&aring; nettet:</span><br />
                              (www.hjemmeside.dk)<br />
                              <input name="webaddress" type="text" class="formInputobjectLong" value="<%=rsPageData("webaddress")%>" size="32" />
                            </td>
                            <td><span class="formLabeltext">Udgiver/Forlag:</span><br />
                                <textarea name="editor" cols="25">
								<%=rsPageData("editor")%></textarea>
                                <span class="formLabeltext"><br />
                                  Email-adresse:</span><br />
                              <input name="editoremail" type="text" class="formInputobjectLong" value="<%=rsPageData("editormail")%>" size="32" />
                              <span class="formLabeltext"><br />
                                Udgiver p&aring; nettet:</span><br />
                              (www.hjemmeside.dk) <br />
                              <input name="editorwww" type="text" class="formInputobjectLong" value="<%=rsPageData("editorwww")%>" size="32" />
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
                                  <option value="">&lt; --- &gt;</option>
                                  <%
IF NOT rsCategories.EOF THEN 
	cursel=CStr(rsCategories.Fields.Item("id").Value & "")
ELSE 
	cursel=""%>
                                  <%END IF
While (NOT rsAllCats.EOF)
%>
                                  <option value="<%=(rsAllCats.Fields.Item("id").Value)%>" <%if (CStr(rsAllCats.Fields.Item("id").Value) = CStr(cursel)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsAllCats.Fields.Item("name").Value)%></option>
                                  <%
  rsAllCats.MoveNext()
Wend
If (rsAllCats.CursorType > 0) Then
  rsAllCats.MoveFirst
Else
  rsAllCats.Requery
End If
%>
                                </select></td>
                            <td><span class="formLabeltext">Stikord:</span><br />
                                <textarea name="keywords" class="formTextobjectLow"><%=rsPageData("keywords")%></textarea>
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
                      <br />
                      <div align="center"><br />
                            <input name="reset" type="reset" class="formbutton" value="Ryd" />
                            <input name="submit" type="submit" class="formSubmitbutton" value="Ret" />
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
<%
rsOrg.Close()
%>
