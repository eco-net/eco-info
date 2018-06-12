<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/admin/functions.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%


username=request("username")
password=request("psw")

IF LEN(username)>0 AND LEN(password)>0 THEN	
	set rs=Server.createobject("ADODB.recordset")
	rs.activeconnection=MM_ecoinfo_STRING
	sql="SELECT * FROM eisys_insiderusers WHERE username='" & replace(username,"'","''") & "' AND password='" & replace(password,"'","''") & "'"
	rs.Source=sql
	rs.Open(sql)
	
	IF NOT rs.EOF THEN
		Session("username")=request("username")
		Session("orgid")=rs.Fields.Item("orgid").value
		rs.close()
		Session("cal")=OrgHas("eical_r_org",Session("orgid"))	
		Session("lib")=OrgHas("eilib_r_org",Session("orgid"))
		Session("art")=OrgHas("eiart_r_org",Session("orgid"))
		
		if Session("orgid")<>"0" then
		response.redirect "/dgs/detail.asp?id=" & Session("orgid")
		else
		response.redirect "/home/index.asp"
		end if
	END IF
		theErr="Brugernavn eller adgangskode er forkert"
ELSE
		theErr="Felter skal udfyldes"
END IF

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="/shared/styles.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="750" border="1" align="center" cellpadding="5" cellspacing="0" bordercolor="#003300" bgcolor="#FFFFFF">
  <tr>
    <th colspan="2" bgcolor="#1979B5" scope="col"><div align="left"><img src="/shared/graphics/logo.gif" width="180" height="33" /></div></th>
  </tr>
  <tr>
    <td valign="top" bgcolor="#E7F1F7"><p align="center" class="contentHeader1">&nbsp;</p>
        <p align="center" class="contentHeader1">Log ind </p>
        <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td scope="col"><div align="left">
                	 <p><%=theErr%></p>
              </div></td>
          </tr>
        </table>
    <p align="center" class="contentHeader1">&nbsp;</p></td>
    <td valign="top" bgcolor="#E7F1F7"><p align="center" class="contentHeader1">&nbsp;</p>
        <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><form id="form1" name="form1" method="post" action="">
                <table width="400" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td colspan="2"><div align="left"><span class="contentHeader1">Log ind p&aring; &Oslash;ko-info</span></div></td>
                  </tr>
                  <tr>
                    <td width="100">brugernavn</td>
                    <td><input name="username" type="text" id="username" /></td>
                  </tr>
                  <tr>
                    <td width="100">adgangskode</td>
                    <td><input name="psw" type="text" id="psw" /></td>
                  </tr>
                  <tr>
                    <td width="100">&nbsp;</td>
                    <td><input type="submit" name="Submit" value="Login" /></td>
                  </tr>
                </table>
            </form></td>
          </tr>
        </table>
    </td>
  </tr>
</table>
</body>
</html>
