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
rsOrg.Source = "SELECT o.id,o.name, o.firstname,o.lastname  FROM eiorg_maindata o LEFT JOIN eilib_r_org l ON o.id=l.orgid  WHERE l.libid=" + Replace(rsOrg__ColParam, "'", "''") + ""
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
rsPageData.Source = "SELECT *  FROM eilib_maindata  WHERE id=" + Replace(rsPageData__ColParam, "'", "''") + ""
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
rsCategories.Source = "SELECT c.id,c.name  FROM eilib_cat_maindata c LEFT JOIN eilib_r_cat o ON c.id=o.catid  WHERE o.libid=" + Replace(rsCategories__ColParam, "'", "''") + ""
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
rsSubjects.Source = "SELECT s.id,s.name  FROM eilib_r_sbj o LEFT JOIN eisbj_maindata s ON o.sbjid=s.id  WHERE o.libid=" + Replace(rsSubjects__ColParam, "'", "''") + ""
rsSubjects.CursorType = 0
rsSubjects.CursorLocation = 2
rsSubjects.LockType = 3
rsSubjects.Open()
rsSubjects_numRows = 0
%>
<%
set rsAllCats = Server.CreateObject("ADODB.Recordset")
rsAllCats.ActiveConnection = MM_ecoinfo_STRING
rsAllCats.Source = "SELECT id,name  FROM eilib_cat_maindata  WHERE siteid=1"
rsAllCats.CursorType = 0
rsAllCats.CursorLocation = 2
rsAllCats.LockType = 3
rsAllCats.Open()
rsAllCats_numRows = 0
%>

<script src="/shared/multiselect.js"></script>
<script src="pub_lib.js"></script>
<script src="/shared/common.js"></script>
<script src="/shared/showeditor.js"></script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Rediger publikation i &Oslash;ko-info</title>
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
                    <a href="#" onClick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">For 
                      MAC brugere.</a></p>
                <p><span class="sidebarHeader">Kategori og emneord</span><br />
                  - er vigtige for at du bliver vist rette sted. Dine valg her er kun vejledende 
                  for vores redakt&oslash;rer der foretager den endelige kategorisering.</p>
                <p><span class="sidebarHeader">Stikord</span><br />
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
                <td class="contentHeader1"> Ret Publikation </td>
              </tr>
              <tr>
                <td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1" /></td>
              </tr>
              <tr>
                <td valign="top"><table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr valign="top">
                        <td> Om Publikationen<br />
                            <span class="formLabeltext">*Titel:</span><br />
                            <input type="text" name="title" value="<%=rsPageData("title")%>" size="32" class="formInputobjectLong" />
                            <br />
                            <span class="formLabeltext">Undertitel</span>: <br />
                            <input type="text" name="subtitle" value="<%=rsPageData("subtitle")%>" size="32" class="formInputobjectLong" />
                            <br />
                            <span class="formLabeltext">*Kategori:</span><br />
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
<option value="<%=(rsAllCats.Fields.Item("id").Value)%>" <%if (CStr(rsAllCats.Fields.Item("id").Value & "") = cursel) then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsAllCats.Fields.Item("name").Value)%></option>
<%
  rsAllCats.MoveNext()
Wend
If (rsAllCats.CursorType > 0) Then
  rsAllCats.MoveFirst
Else
  rsAllCats.Requery
End If
%>
</select>

                            <br />
                          <span class="formLabeltext">*Forfatter:</span><br />
                          <input type="text" name="author" value="<%=rsPageData("author")%>" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">ISBN:</span><br />
                          <input type="text" name="isbn" value="<%=rsPageData("isbn")%>" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">Udgivelses&aring;r:</span><br />
                          <input type="text" name="year" value="<%=rsPageData("year")%>" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">Sprog:</span><br />
                          <input type="text" name="language" value="<%=rsPageData("language")%>" size="32" class="formInputobjectLong" />
                          <br />
                          <span class="formLabeltext">P&aring; nettet:</span><br />
                          <input type="text" name="webaddress" value="<%=rsPageData("webaddress")%>" size="32" class="formInputobjectLong" />
                        </td>
                        <td> Om udgiveren<br />
                            <span class="formLabeltext">Udgiver/Forlag:</span><br />
                            <textarea name="editor" cols="32" class="formTextobjectLow"><%=rsPageData("editor")%></textarea>
                            <span class="formLabeltext"><br />
                              Email-adresse:</span><br />
                          <input type="text" name="editoremail" value="<%=rsPageData("editoremail")%>" size="32" class="formInputobjectLong" />
                          <span class="formLabeltext"><br />
                            P&aring; nettet:</span><br />
                          <input type="text" name="editorwww" value="<%=rsPageData("editorwww")%>" size="32" class="formInputobjectLong" /></td>
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
                      <textarea name="shortdescription" cols="50" rows="5" class="formTextobjectLow">
					  <%=rsPageData("shortdescription")%></textarea>
                      <div align="center"></div></td>
                      <td><span class="formLabeltext">*Resum&eacute;:</span><br />
                      <textarea name="description" cols="50" rows="5" class="formTextobjectBig">
					  <%=rsPageData("description")%></textarea>
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
                      <select name="selSbj" class="formTextobjectLow" size="5" ondblclick="removeValue(this.form.selSbj);" multiple>
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
                        <textarea name="keywords" class="formTextobjectLow"><%=rsPageData("keywords")%></textarea></td>
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
                  <input name="submit" type="submit" class="formSubmitbutton" value="Ret" />
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
<%
rsOrg.Close()
%>
