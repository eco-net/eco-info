<%
IF LEN(request("step"))=0 OR request("step")="1" THEN
	step=1
ELSE
	step=CINT(request("step"))
END IF

IF step=3 THEN

	theSubj="Tilmelding fra nettet"
	theBody=request("tfname") & " " & request("tlname") & " ønsker at modtage følgende:" & VbCrLf & VbCrLf
	IF LEN(request("cmail"))>0 THEN theBody=theBody & "- Nyhedsmail" & VbCrLf
	IF LEN(request("cmag"))>0 THEN theBody=theBody & "- kampagnemateriale fra Øko-net" & VbCrLf
	IF LEN(request("cmember"))>0 THEN theBody=theBody & "- Medlemsskab" & VbCrLf
	theBody=theBody & VbCrLf & VbCrLf & VbCrLf
	theBody=theBody & "Kopier nedenstående til oprettelse i Filemaker:" & VbCrLf
	theBody=theBody & "Fornavn:" & request("tfname") & ";Efternavn:" & request("tlname") & ";Mail:" &_
		request("tmail") & ";Gade:" & request("tstreet") & ";Postnr:" & request("tzip") & ";By:" &_
		request("tCity") & ";"
	IF LEN(request("cmail"))>0 THEN theBody=theBody & "Nyhedsmail;" & VbCrLf
	IF LEN(request("cmag"))>0 THEN theBody=theBody & "Nyhedsblad;" & VbCrLf
	IF LEN(request("cmember"))>0 THEN theBody=theBody & "Medlemsskab;" & VbCrLf & VbCrLf
	theBody=theBody & VbCrLf & VbCrLf & "Afsendt den " & FormatDateTime(Date(),vbLongDate) & " kl. " & FormatDateTime(Now(),vbShortTime)

	'SendMail "eco-net@post8.tele.dk","info@eco-net.dk",theSubj,theBody
	'SendCDOMail "eco-net@post8.tele.dk","eco-net@eco-net.dk",theSubj,theBody
	SendCDOMail "eco-net@eco-net.dk","eco-net@eco-net.dk",theSubj,theBody
END IF

SUB SendCDOMail(thefrom,theto,theSubj,theBody)
Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = theSubj
objMessage.From = thefrom 
objMessage.To = theto
objMessage.TextBody = theBody
objMessage.Send
END SUB

SUB SendMail(theTo,theFrom,theSubject,theBody)
	Set msg = Server.CreateObject("JMail.Message")
	msg.From = theFrom
	msg.FromName = "Øko-net"
	msg.Charset = "iso-8859-1"
	msg.AddRecipient theTo
	msg.Subject = theSubject
	msg.Body = theBody
	msg.Send("mail.eco-info.dk" )
	Set msg = nothing
END SUB

%>

<html>
<!-- #BeginTemplate "/Templates/noheader_okonet.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>&Oslash;ko-info - guide til alt om &oslash;kologi, natur, milj&oslash; og b&aelig;redygtighed</title>
<script language="JavaScript">
function gonext(theform){
	var step=theform.step.value;
	var valid=false;
	switch(step){
		case '1':
			if(!theform.cmail.checked && !theform.cmag.checked && !theform.cmember.checked)alert('Afkryds de(n) ønskede service.');else valid=true;
			break;
		case '2':
			if(!theform.tfname.value.length || !theform.tlname.value.length)alert('Du har glemt at indtaste dit navn.');
			else if(theform.cmail.value==1 && !theform.tmail.value.length)alert('Du har glemt at indtaste din email adresse.');
			else if((theform.cmag.value.length || theform.cmember.value.length) && !theform.tstreet.value.length)alert('Du har glemt at indtaste din adresse.');
			else valid=confirm('Dine oplysninger sendes nu til Øko-net\'s sekretariat.\nKontrollér venligst at de er korrekte.\n\nKlik Annuller for at ændre noget.');
			break;
		case '3':
			window.close();
			break;
	}
	if(valid){theform.step.value=parseInt(step)+1;theform.submit();}
}

function goprev(theform){
	theform.step.value-=1;theform.submit();
}
</script>
<!-- #EndEditable --> 
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="7" marginwidth="0" marginheight="7">
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" name="Pagetable">
<tr> 
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
<td width="100%" valign="top"> 
<!-- START HEADER -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" name="Header">
<tr> 
<td valign="top" rowspan="3" width="180" heigth="33"><img src="/shared/graphics/logo_okonet.gif" width="180" height="33"></td>
<td height="17"><br>
</td>
</tr>
<tr>
<td valign="top" width="99%" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr>
<td height="16"><br>
</td>
</tr>
</table>
<!-- END HEADER -->

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="100%" valign="top" name="maincontent">
<table border="0" cellspacing="0" cellpadding="10" width="100%" name="Contentarea">
<tr>
<td valign="top">
<!-- #BeginEditable "maincontent" -->
<table border="0" cellspacing="0" cellpadding="10" width="100%">
<tr>
<td>
<form name="form1" method="post" action="">
<%IF step=1 THEN%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td height="12" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="sidebarHeaderBox" height="20">
<tr>
<td class="faktaboksheader" width="33%">1. V&aelig;lg  services</td>
<td width="33%">2. Personlige data</td>
<td width="34%">3. Bekr&aelig;ftelse</td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign="top"><span class="contentHeader1">&Oslash;nsker du at ...</span><br>
<br>
<br>
<input type="checkbox" name="cmail" value="1" <%IF LEN(request("cmail"))>0 THEN response.write "checked"%>>
&nbsp;&nbsp;&nbsp;
Modtage vores nyhedsmail.<br>
<br>
<input type="checkbox" name="cmag" value="1" <%IF LEN(request("cmag"))>0 THEN response.write "checked"%>>
&nbsp;&nbsp;&nbsp;
Modtage eksempler på seneste nyt kampagnemateriale fra Øko-net.<br>
<br>
<br>
</td>
<td>&nbsp;</td>
</tr>
<tr>
<td colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="50%" align="right">&nbsp;</td>
<td width="50%">
<input type="button" name="Button" value="N&aelig;ste &gt;&gt;" class="formbutton" onClick="gonext(this.form)">
<input type="hidden" name="tfname" value="<%= Request("tfname") %>">
<input type="hidden" name="tlname" value="<%= Request("tlname") %>">
<input type="hidden" name="tmail" value="<%= Request("tmail") %>">
<input type="hidden" name="tstreet" value="<%= Request("tstreet") %>">
<input type="hidden" name="tzip" value="<%= Request("tzip") %>">
<input type="hidden" name="tcity" value="<%= Request("tcity") %>">

</td>
</tr>
</table>
</td>
</tr>
</table>
<%ELSEIF step=2 THEN%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td height="12" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="sidebarHeaderBox" height="20">
<tr>
<td width="33%">1. V&aelig;lg services</td>
<td class="faktaboksheader" width="33%">2. Personlige data</td>
<td width="34%">3. Bekr&aelig;ftelse</td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign="top"><span class="contentHeader1">Indtast personlige oplysninger ...</span><br>
<br>
<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td class="formLabeltext">Fornavn</td>
<td>
<input type="text" name="tfname" class="formInputobject" value="<%= Request("tfname") %>">
</td>
</tr>
<tr>
<td class="formLabeltext">Efternavn</td>
<td>
<input type="text" name="tlname" class="formInputobject" value="<%= Request("tlname") %>">
</td>
</tr>
<tr>
<td class="formLabeltext">Gade / Vej</td>
<td>
<input type="text" name="tstreet" class="formInputobjectLong" value="<%= Request("tstreet") %>">
</td>
</tr>
<tr>
<td class="formLabeltext">Postnr og By</td>
<td>
<input type="text" name="tzip" class="formInputobject" value="<%= Request("tzip") %>">
<input type="text" name="tcity" class="formInputobject" value="<%= Request("tcity") %>">
</td>
</tr>
<tr>
<td class="formLabeltext" nowrap>Email adresse&nbsp;&nbsp;&nbsp;</td>
<td>
<input type="text" name="tmail" class="formInputobjectLong" value="<%= Request("tmail") %>">
</td>
</tr>
</table>
<br>
<br>
</td>
<td>&nbsp;</td>
</tr>
<tr>
<td colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="50%" align="right">
<input type="button" name="Button" value="&lt;&lt; Tilbage" class="formbutton" onClick="goprev(this.form)">
</td>
<td width="50%">
<input type="button" name="Button" value="N&aelig;ste &gt;&gt;" class="formbutton"  onClick="gonext(this.form)">
<input type="hidden" name="cmail" value="<%= Request("cmail") %>">
<input type="hidden" name="cmag" value="<%= Request("cmag") %>">
<input type="hidden" name="cmember" value="<%= Request("cmember") %>">
</td>
</tr>
</table>
</td>
</tr>
</table>
<%ELSEIF step=3 THEN%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td height="12" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="sidebarHeaderBox" height="20">
<tr>
<td width="33%">1. V&aelig;lg services</td>
<td width="33%">2. Personlige data</td>
<td width="34%" class="faktaboksheader">3. Bekr&aelig;ftelse</td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign="top"><span class="contentHeader1">Tak for din henvendelse ...</span><br>
<br>
<br>
<%IF LEN(request("cmag"))>0 OR LEN(request("cmember"))>0 THEN%>
Du vil modtage det ønskede materiale med posten en af de nærmeste dage.
<%IF LEN(request("cmember"))>0 THEN%>
<br>
<br>
Betal dit medlemsskab med det samme via ewire:<br><br>
<script type='text/javascript' language='javascript' src='http://www.ewire.dk/includes/quickservice_functions.js.asp'></script>
<div align="center">
<a style="cursor:pointer" onClick="javascript:open_ewire_email('eco-net@eco-net.dk','%D8ko-net+','St%F8ttebel%F8b','','200,00','24','1');">
<img src='http://www.ewire.dk/images/icons/ewire_donation.gif' border='0'></a>
</div>
<%END IF%>
<%IF LEN(request("cmail"))>0 THEN%>
<br><br>
<%END IF%>
<%END IF%>
<%IF LEN(request("cmail"))>0 THEN%>
Du vil fremover modtage vores nyhedsmails.
<%END IF%>
<br>
<br>
<%IF NOT LEN(request("cmember"))>0 THEN%>
<strong>STØT ØKO-NET</strong><br>
Støt Øko-net med 50 kr. om året.<br>
Overfør pengene nemt og let via e-mail på ewire.
<br>
<br>
<script type='text/javascript' language='javascript' src='http://www.ewire.dk/includes/quickservice_functions.js.asp'></script>
<div align="center">
<a style="cursor:pointer" onClick="javascript:open_ewire_email('eco-net@eco-net.dk','%D8ko-net+','St%F8ttebel%F8b','','50,00','24','1');">
<img src='http://www.ewire.dk/images/icons/ewire_donation.gif' border='0'></a>
</div>
<br>
<br>
<%END IF%>

</td>
<td>&nbsp;</td>
</tr>
<tr>
<td colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="50%" align="right">&nbsp;
</td>
<td width="50%">
<input type="hidden" name="tfname" value="<%= Request("tfname") %>">
<input type="hidden" name="tlname" value="<%= Request("tlname") %>">
<input type="hidden" name="tmail" value="<%= Request("tmail") %>">
<input type="hidden" name="tstreet" value="<%= Request("tstreet") %>">
<input type="hidden" name="tzip" value="<%= Request("tzip") %>">
<input type="hidden" name="tcity" value="<%= Request("tcity") %>">
<input type="hidden" name="cmail" value="<%= Request("cmail") %>">
<input type="hidden" name="cmag" value="<%= Request("cmag") %>">
<input type="hidden" name="cmember" value="<%= Request("cmember") %>">
</td>
</tr>
</table>
</td>
</tr>
</table>
<%END IF%>
<input type="hidden" name="step" value="<%= step %>">
</form>
</td>
</tr>
<tr>
<td align="right">
<a href="javascript:window.close();">Luk vindue</a>
</td>
</tr>
</table>
<!-- #EndEditable -->
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
</tr>
<tr height="1"> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="3"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
</body>
<!-- #EndTemplate -->
</html>
<!--include virtual="/shared/log.asp"-->
