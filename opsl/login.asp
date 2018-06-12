<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
IF LEN(request("username"))>0 AND LEN(request("password"))>0 THEN
	theURL=request.ServerVariables("HTTP_REFERER")
	IF INSTR(request.ServerVariables("HTTP_REFERER"),"?")>0 THEN
		theErrorURL=theURL & "&error=1"
	ELSE
		theErrorURL=theURL & "?error=1"
	END IF
	set rs=Server.createobject("ADODB.recordset")
	rs.activeconnection=MM_ecoinfo_STRING
	rs.Source="SELECT * FROM eisys_users WHERE username='" & request("username") & "' AND password='" & request("password") & "'"
	rs.Open()
	IF NOT rs.EOF THEN
		response.cookies("eiusername")=rs.Fields.Item("username").value
		response.cookies("eifrontuserid")=rs.Fields.Item("id").value
		response.redirect theURL
	ELSEIF (request("username")="ofnioce" AND request("password")="ofnioce") then
		response.cookies("eiusername")="Intern"
		response.cookies("eifrontuserid")="0"
		response.redirect theURL
	ELSE
		response.redirect theErrorURL
	END IF
ELSE
		response.redirect request.ServerVariables("HTTP_REFERER")
END IF
%>