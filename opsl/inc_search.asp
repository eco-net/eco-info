<script src="/log/optionfiles/selamtkomm.js"></script>
<script language="JavaScript">
function submitSearch(theform) {
	if(theform.opslcat.selectedIndex>0)theform.opslcatname.value=stripit(theform.opslcat.options[theform.opslcat.selectedIndex].text);
	if(theform.opslkomm.selectedIndex>0)theform.opslkommname.value=theform.opslkomm.options[theform.opslkomm.selectedIndex].text;
	if(theform.opslamt.selectedIndex>0)theform.opslamtname.value=theform.opslamt.options[theform.opslamt.selectedIndex].text;
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
<input type="hidden" name="opsltimename" value="">
<input type="hidden" name="opslcatname" value="">
<input type="hidden" name="opslsbjname" value="">
<input type="hidden" name="opslkommname" value="">
<input type="hidden" name="opslamtname" value="">
<input type="hidden" name="sektion" value="ok">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr bgcolor="#FAFAF4"> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;VIS OPSLAG</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><span class="formLabeltext">I kategorien</span>...<br>
<select name="opslcat" class="formInputobject">
<script src="/log/optionfiles/opslcat_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i regionen </span>...<br>
<select name="opslamt" class="formInputobject" onChange="populateSel(this.form.opslkomm,this.form.opslamt.options[this.form.opslamt.selectedIndex].value)">
<script src="/log/optionfiles/amter_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i kommunen</span>...<br>
<select name="opslkomm" class="formInputobject">
<script src="/log/optionfiles/kommuner_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) med teksten</span>...<br>
<input type="text" name="opsltext" class="formInputobject">
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
