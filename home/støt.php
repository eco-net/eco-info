<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>St&oslash;t &Oslash;ko-net</title>



<?
$destinatar = "eco-net@eco-net.dk";
$subiect = "Støttebidrag - ";
?>
<script language="javascript">
function isEmail(str)
{
	var supported = 0;
	if (window.RegExp)
	{
		var tempStr = "a";
		var tempReg = new RegExp(tempStr);
		if (tempReg.test(tempStr)) supported = 1;
	}
	if (!supported) 
		return (str.indexOf(".") > 2) && (str.indexOf("@") > 0);
	var r1 = new RegExp("(@.*@)|(\\.\\.)|(@\\.)|(^\\.)");
	var r2 = new RegExp("^.+\\@(\\[?)[a-zA-Z0-9\\-\\.]+\\.([a-zA-Z]{2,3}|[0-9]{1,3})(\\]?)$");
	return (!r1.test(str) && r2.test(str));
}

	function enableButton()
	{
		if(sendMail.nume.value.length < 3)
		{
			sendMail.trimite.disabled = true;
		}
		/*else if(sendMail.telefon.value.length == 0)
		{
			sendMail.trimite.disabled = true;
		}*/
		
		else if (sendMail.email.value.length == 4)
		{
			sendMail.trimite.disabled = true;
		}
		
		else if(sendMail.subiect.value.length == 2)
		{
			sendMail.trimite.disabled = true;
		}
		/*else if(sendMail.optiuni.value == "Alege")
		{
			sendMail.trimite.disabled = true;
		}*/
		else if(sendMail.mesaj.value.length < 2)
		{
			sendMail.trimite.disabled = true;
		}
		else
		{
			sendMail.trimite.disabled = false;
		}
	}
</script>









<style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 13px;
	font-weight: bold;
}
.style2 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px; }
.style6 {
	font-size: 9px;
	color: #999999;
}
-->
</style>
</head>

<body>

 <?

if(!isset($_POST['beløb']) && !isset($_POST['fornavn']) && !isset($_POST['efternavn']) && !isset($_POST['gade']) && !isset($_POST['postnr'])&& !isset($_POST['by'])&& !isset($_POST['tel'])&& !isset($_POST['mob'])&& !isset($_POST['email'])&& !isset($_POST['RadioGroup1'])&& !isset($_POST['RadioGroup2']))
{
//Primul ecran, unde se completeaza formularul
?>

<form name="sendMail" action="" method="post">
<table width="541" border="0" cellpadding="0" cellspacing="6">
  <tr>
    <td><img src="http://web.eco-net.dk/shared/graphics/logo4_m.gif" width="529" height="31" /></td>
  </tr>
  <tr>
    <td><p class="style1"><br />St&oslash;t &Oslash;ko-net </p>
      <p class="style2">Som folkeoplysningsorganisation er vi &oslash;konomisk afh&aelig;ngig af vores medlemmer og folks (privatpersoner/institutioner) st&oslash;ttebidrag for at lave vores arrangementer, seminarer og projekter. Du v&aelig;lger selv st&oslash;ttebidragets st&oslash;rrelse - minimum 50 Kr. Til geng&aelig;ld viser du aktivt din holding til &oslash;kologi og b&aelig;redygtig udvikling. Du giver os muglighed for at sk&aelig;rpe fokus og pr&aelig;ge den politiske dagsorden i endnu h&oslash;jere grad. </p>
      <p class="style2">1. Giv et bidrag </p>
	
	
	
	  
      <table width="245" border="0" cellpadding="6">
        <tr>
          <td width="37" class="style2">Bel&oslash;b:</td>
          <td width="178" class="style2"><input name="beløb" type="text" size="8" maxlength="6" />
            Kr. <span class="style6"> (Min. 50 Kr.)</span> </td>
        </tr>
      </table>
      <p class="style2">2. Angiv personoplysninger </p>
      <table width="21%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td>




<table width="99%" border="0" cellspacing="0" cellpadding="6">
          <tr> 
            <td width="16%" class="style2">Fornavn:</td>
            <td width="84%" class="style2"><input name="fornavn" type="text" maxlength="20"></td>
          </tr>
          <tr>
            <td class="style2">Efternavn:</td>
            <td class="style2"><input name="efternavn" type="text" maxlength="20" /></td>
          </tr>
          
          <tr> 
            <td class="style2">Gade/Vej:</td>
            <td class="style2"><input name="gade" type="text" maxlength="30"></td>
          </tr>
          <tr> 
            <td class="style2">Postnr.:</td>
            <td class="style2"><input name="postnr" type="text" maxlength="8"></td>
          </tr>
          <tr> 
            <td class="style2">By:</td>
            <td class="style2"><input name="by" type="text" maxlength="20"></td>
          </tr>
          <tr>
            <td class="style2">Telefonnummer:</td>
            <td class="style2"><input name="tel" type="text" maxlength="20" /></td>
          </tr>
          <tr>
            <td class="style2">Mobilnummer: <span class="style6">(valgfri)</span> </td>
            <td class="style2"><input name="mob" type="text" maxlength="20" /></td>
          </tr>
          <tr>
            <td class="style2">E-mail:</td>
            <td class="style2"><input name="email" type="text" maxlength="20"></td>
          </tr>
        </table> </td>
  </tr>
</table>
    <p class="style2">3. V&aelig;lg betalingsmetode</p>
    <p class="style2">Jeg &oslash;nsker at betale      </p>
    <p>
      <label>
        
        <span class="style2">
        <input type="radio" name="RadioGroup2" value="giro" />
        Via giro</span></label>
      <span class="style2"><br />
      <label>
        <input type="radio" name="RadioGroup2" value="PBS" />
        Via PBS</label>
      </span>
      <label></label>
    </p>    </td>
  </tr>
  <tr>
    <td width="100%" height="192"><p class="style2">Hvis du &oslash;nsker at betale dit  st&oslash;ttebidrag til &Oslash;ko-net via betalingsservice, skal du klikke p&aring; BS  betalingsservice-logoet. Der kommer et pop-up-vindue, hvor du genindtaster dine  personlige oplysninger. I medlemsnummerfeltet skal du indtaste dit  telefonnummer. Herefter trykker du <strong>Send</strong> til Betalingsservice, og vender tilbage til denne side.</p>
      <span class="style2"><strong>Vigtigt:</strong> Husk at trykke tilmeld her p&aring; siden, s&aring; &Oslash;ko-net ogs&aring; modtager dine oplysninger. </span>
      <p align="center" class="style2"><a href="#" onclick="window.open('https://www.betalingsservice.dk/BS/?id=0&pbs=6bfbc0511c12c1cab6024a28f0de59ad&dbnr=&dbgr=&navn=&adr=&postby=&knmin=&knmax=','BSTilmeld','resizable=no,toolbar=no,scrollbars=no,menubar=no,location=no,status=yes,width=375,height=500');returnfalse;"><img src="http://web.eco-net.dk/shared/graphics/Betalingsservice.gif" width="200" height="28" border="0" /></a></p>
      <p align="left" class="style2">Betal dit st&oslash;ttebidrag med betalingsservice.  Det sikrer, at dit bidrag automatisk bliver betalt til tiden og du sparer  gebyrerne p&aring; posthuset eller i banken.</p></td>
    </tr>
  <tr>
    <td height="46"><div align="center"><span class="style2">
      <input name="button" type="submit" class="style2" value="Tilmeld" />
    </span></div></td>
  </tr>
</table>

</form>


 <?

}

else

{
	$beløb = $_POST['beløb'];
	$fornavn = $_POST['fornavn'];
	$efternavn = $_POST['efternavn'];
	

	$gade = $_POST['gade'];
	
	
	$postnr = $_POST['postnr'];
	$by = $_POST['by'];
	$tel = $_POST['tel'];
	$mob = $_POST['mob'];
	$email = $_POST['email'];
	$RadioGroup1 = $_POST['RadioGroup1'];
	$RadioGroup2 = $_POST['RadioGroup2'];

	$ip = $_SERVER['REMOTE_ADDR'];

	

	$headers  = "MIME-Version: 1.0\r\n";

	$headers .= "Content-type: text/html; charset=iso-8859-1\r\n";

	$headers .= "Støttebidrag fra nettet\r\n";

	$subiect .= "fra nettet\r\n";

	$mesaj = nl2br(str_replace("\\","",$email));

	
	//trimite mailul, afiseaza mesajul de confirmare
	$mesajConfirmare = "<b>Dine oplysninger blev sent til Øko-net<br>Tak for dit støttebidrag</b>\r\n";
	echo $mesajConfirmare;
	
	

	$mesaj .= "<hr><br><b><u>Støttebidrag fra:</u></b><br>Fornavn:$fornavn;Efternavn:$efternavn;Mail:$email;Gade:$gade;Postnr:$postnr;By:$by;Støttebidrag:$beløb;Telefonnummer:$tel;Mobilnummer:$mob;Betale via:$RadioGroup2";

	mail($destinatar, $subiect, $mesaj, $headers);

}

?>
	 
	 

</body>
</html>
