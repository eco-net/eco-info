<script src="/log/optionfiles/selamtkomm.js"></script>
<script language="JavaScript">
function submitSearch(theform) {
	if(theform.dgscat.selectedIndex>0)theform.dgscatname.value=stripit(theform.dgscat.options[theform.dgscat.selectedIndex].text);
	if(theform.dgssbj.selectedIndex>0)theform.dgssbjname.value=stripit(theform.dgssbj.options[theform.dgssbj.selectedIndex].text);
	if(theform.dgskomm.selectedIndex>0)theform.dgskommname.value=theform.dgskomm.options[theform.dgskomm.selectedIndex].text;
	if(theform.dgsamt.selectedIndex>0)theform.dgsamtname.value=theform.dgsamt.options[theform.dgsamt.selectedIndex].text;
	return true
}
</script>

<%IF LEN(recInfo)>0 THEN%>
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;AKTUEL S&Oslash;GNING</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr> 
<td>
<%=searchDescr%>
</td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
<%END IF%>
<table width="180" border="0" cellspacing="0" cellpadding="0">
<form name="theform" method="post" action="list.asp" onSubmit="return submitSearch(document.theform);">
<input type="hidden" name="dgscatname" value="">
<input type="hidden" name="dgssbjname" value="">
<input type="hidden" name="dgskommname" value="">
<input type="hidden" name="dgsamtname" value="">
<input type="hidden" name="sektion" value="dgs">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;VIS ORGANISATIONER</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><span class="formLabeltext">I kategorien</span>...<br>
<select name="dgscat" class="formInputobject">
<script src="/log/optionfiles/dgscat_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) under emnet</span>...<br>
<select name="dgssbj" class="formInputobject">
<script src="/log/optionfiles/sbj_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i regionen </span>...<br>
<select name="dgsamt" class="formInputobject" onChange="populateSel(this.form.dgskomm,this.form.dgsamt.options[this.form.dgsamt.selectedIndex].value)">
<script src="/log/optionfiles/amter_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i kommunen</span>...<br>
<select name="dgskomm" class="formInputobject">
<script src="/log/optionfiles/kommuner_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) med teksten</span>...<br>
<input type="text" name="dgstext" class="formInputobject">
</td>
</tr>
<tr> 
<td align="center"> <br>
<input type="submit" value="S&oslash;g" class="formSubmitbutton">
</td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</form>
</table>
