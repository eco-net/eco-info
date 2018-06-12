<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->

<%
SUB DoSave(thefilename,act)
	IF act=1 THEN 'Insert
		DoCreate thefilename
	ELSEIF act=4 THEN 'Delete
		'DeleteFileInDB "eiopsl_maindata","billede_filnavn","id",request("id")
		theSQL="Delete From eiopsl_maindata WHERE id=" & request("id")
		execCommand theSQL
		'**** Showing confirmation to user
		response.redirect("confirmation.asp?done=4")
	END IF
END SUB

SUB DoCreate(thefilename)

	DIM userpw,userID
	'IF there's username and password, the user is created
	
	IF LEN(UploadRequest.Item("postnr").Item("Value"))>0 THEN
		pnr=UploadRequest.Item("postnr").Item("Value") + 0
	ELSE
		pnr=0
	END IF	
	IF LEN(UploadRequest.Item("username").Item("Value"))>0 AND LEN(UploadRequest.Item("userid").Item("Value"))=0 THEN
		userpw=GetUniquepw(UploadRequest.Item("username").Item("Value"),UploadRequest.Item("password").Item("Value"))
		theSQL="INSERT INTO eisys_users (name,username,password,zipcode,phone,emailaddress) VALUES (" &_
		"'" & replace(UploadRequest.Item("indsender").Item("Value") & "","'","''") & "'" & "," &_
		"'" & replace(UploadRequest.Item("username").Item("Value") & "","'","''") & "'" & "," &_
		"'" & userpw & "'" & "," &_
		pnr & "," &_
		"'" & UploadRequest.Item("telefon").Item("Value") & "'" & "," &_
		"'" & UploadRequest.Item("email").Item("Value") & "'" & ")"
		userID=InsertRec(theSQL,"eisys_users","id","")	
	ELSEIF LEN(UploadRequest.Item("userid").Item("Value"))>0 THEN
		userID=UploadRequest.Item("userid").Item("Value")	
		userpw=""
	ELSE
		userID=0	
		userpw=""
	END IF
	
	
	
	
	'FormatDateDB(UploadRequest.Item("exitdato").Item("Value")) & "," &_
	theSQL="INSERT INTO eiopsl_maindata(" &_
		"title,indsender,telefon,email,cat_id,postnr,kort_beskrivelse,lang_beskrivelse,billede_filnavn,billede_width,billede_height,userid,siteid" &_
		") VALUES (" &_
		"'" & replace(UploadRequest.Item("titel").Item("Value") & "","'","''") & "'" & "," &_
		"'" & replace(UploadRequest.Item("indsender").Item("Value") & "","'","''") & "'" & "," &_
		"'" & UploadRequest.Item("telefon").Item("Value") & "'" & "," &_
		"'" & UploadRequest.Item("email").Item("Value") & "'" & "," &_
		UploadRequest.Item("cat_id").Item("Value") & "," &_
		pnr & "," &_
		"'" & replace(UploadRequest.Item("kort_beskrivelse").Item("Value") & "","'","''") & "'" & "," &_
		"'" & replace(UploadRequest.Item("lang_beskrivelse").Item("Value") & "","'","''") & "'" & "," &_
		"'" & thefilename & "'," &_
		UploadRequest.Item("pwidth").Item("Value") & "," &_
		UploadRequest.Item("pheight").Item("Value") & "," &_
		userid & ",1)"
	
	execCommand theSQL

	'**** Showing confirmation to user
	IF userpw="" THEN
		response.redirect("confirmation.asp?done=1")
	ELSE
		response.cookies("eiusername")=UploadRequest.Item("username").Item("Value")
		response.cookies("eifrontuserid")=userID
		response.redirect("confirmation.asp?done=1&userpw=" & userpw)
	END IF
END SUB


%>