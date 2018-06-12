<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#FAFAF4">
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;BRUG OPSLAGTAVLEN</td>
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
<form name="form1" method="post" action="new.asp">
Klik nedenfor for at oprette et opslag.
<p>Inden du g&aring;r igang vil det v&aelig;re en god id&eacute; at g&oslash;re 
dig klart hvad du vil skrive, samt evt. finde et billede frem. Billeder kan maksimalt 
v&aelig;re 500 x 500 pixel.</p>
<div align="center">
<input type="submit" name="Submit" value="Nyt Opslag" class="formbutton">
</div>
</form>
<form name="form1" method="post" action="login.asp">
<span class="SidebarHeader">Log ind</span><br>
<%IF LEN(request.cookies("eiusername"))=0 THEN%>
Log ind for at slette et opslag du har lavet, eller slippe for at indtaste dine personlige oplysninger en gang til når du laver nye opslag.<br><br>
<span class="formLabeltext">
Brugernavn:</span>
<input type="text" name="username" class="formInputobject">
<span class="formLabeltext"><br>
Adgangskode:</span><br>
<input type="text" name="password" class="formInputobject">
<div align="center"><br>
<input type="submit" name="Submit" value="Log ind" class="formSubmitbutton" onClick="if(this.form.username.length==0 || this.form.password.length==0)alert('Begge felter skal udfyldes.');else{this.form.action='login.asp';this.form.submit();}">
</div>
<%ELSE%>
Du er logget ind som <b><%=request.cookies("eiusername")%></b>.<br>
<div align="center"><br>
<%IF LEN(request.cookies("eiuserid"))>0 THEN%>
<input type="submit" name="Submit" value="Slet Opslag" class="formbutton" onClick="this.form.action='delete.asp';this.form.submit();">
<% end if %>
</div>

<%END IF%>
</form>
</td>
</tr>
<tr>
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
