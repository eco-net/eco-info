<!--#include virtual="/connections/ecoinfo.asp" -->

<%
function picasaImgSize(imgsrc,w)
i=1
imgname=Right(imgsrc,i) 
while InStr(imgname,"/")=0 
i=i+1
imgname=Right(imgsrc,i) 
wend
imgsrc=Left(imgsrc,Len(imgsrc)-i)
if Right(imgsrc,5)="/s144" then 'hvis thumbnail er kopieret
imgsrc=Left(imgsrc,Len(imgsrc)-5)
end if
if Right(imgsrc,5)="/s400" then 'hvis s400 er kopieret
imgsrc=Left(imgsrc,Len(imgsrc)-5)
end if
if Right(imgsrc,5)="/s512" then 'hvis s512 er kopieret
imgsrc=Left(imgsrc,Len(imgsrc)-5)
end if
if Right(imgsrc,5)="/s576" then 'hvis s576 er kopieret
imgsrc=Left(imgsrc,Len(imgsrc)-5)
end if
if Right(imgsrc,5)="/s640" then 'hvis s640 er kopieret
imgsrc=Left(imgsrc,Len(imgsrc)-5)
end if
if Right(imgsrc,5)="/s720" then 'hvis s720 er kopieret
imgsrc=Left(imgsrc,Len(imgsrc)-5)
end if
if Right(imgsrc,5)="/s800" then 'hvis s800 er kopieret
imgsrc=Left(imgsrc,Len(imgsrc)-5)
end if
imgsrc=imgsrc & w & imgname
picasaImgSize=imgsrc
end function

FUNCTION ReadFile(filename)
	DIM fs,ts,theText
	theText="file created"
	filename=Server.MapPath(filename)
	SET fs = CreateObject("Scripting.FileSystemObject")
	IF fs.FileExists(filename)=false THEN
		SET ts = fs.CreateTextFile(filename)
		ts.Write(theText)
	ELSE		
		Set ts=fs.OpenTextFile(filename)
		theText=ts.ReadAll()
	END IF
	fs=""
	ts=""
	ReadFile=theText
END FUNCTION

SUB WriteFile(theText,filename)

	DIM fs,ts
	filename=server.MapPath(filename)
	
	SET fs = createobject("Scripting.FileSystemObject")
	Set ts=fs.CreateTextFile(filename,true)
	ts.Write(theText)
	fs=""
	ts=""
	
END SUB

SUB DeleteFile(filename)
	DIM fso,fo
	filename=Server.MapPath(filename)
	Set fso = CreateObject("Scripting.FileSystemObject")
	IF fso.FileExists(filename)=true THEN
		SET fo=fso.GetFile(filename)
		fo.Delete(true)
		fo=""
	END IF
	fso=""
END SUB

SUB CreateFolder(foldername)
	DIM fso
	foldername=Server.MapPath(foldername)
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.CreateFolder(foldername)
	fso=""
END SUB

SUB DeleteFolder(foldername)
	DIM fso
	foldername=Server.MapPath(foldername)
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFolder(foldername)
	fso=""
END SUB

SUB SendMail(theTo,theFrom,theSubject,theBody)
	Set msg = Server.CreateOBject( "JMail.Message" )
	msg.From = theFrom
	msg.FromName = "Øko-info"
	msg.Charset = "iso-8859-1"
	msg.AddRecipient theTo
	msg.Subject = theSubject
	msg.Body = theBody
	msg.Send("194.255.126.227" )
	Set msg = nothing
END SUB

SUB SendCDOMail(theto,thefrom,theSubj,theBody)
Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = theSubj
objMessage.From = thefrom 
objMessage.To = theto
'objMessage.TextBody = theBody
objMessage.HTMLBody = theBody
objMessage.Send
set objMessage=nothing
END SUB

SUB DeleteFileInDB(theTable,theFileField,theidField,theID)
	DIM theRS, fm
	SET theRS = Server.CreateObject("ADODB.recordset")
	theRS.ActiveConnection=MM_ecoinfo_STRING
	theRS.Source="SELECT " & theFileField & " FROM " & theTable & " WHERE " & theIDField & "=" & theID
	theRS.Open()
	DeleteFile theRS.Fields.Item(theFileField).value
	theRS.Close()
END SUB

FUNCTION GetUniquePW(username,password)
	DIM theRS,x,tpassword
	SET theRS = Server.CreateObject("ADODB.recordset")
	theRS.ActiveConnection=MM_ecoinfo_STRING
	theRS.Source="SELECT id FROM eisys_users WHERE username='" & username & "' AND password='" & password & "'"
	theRS.Open()
	x=0
	WHILE NOT theRS.EOF
		theRS.Close()
		x=x+1
		tpassword=password & CSTR(x)
		theRS.Source="SELECT id FROM eisys_users WHERE username='" & username & "' AND password='" & tpassword & "'"
		theRS.Open()
	WEND
	theRS.Close()
	IF x>0 THEN
		GetUniquePW=password & CSTR(x)
	ELSE
		GetUniquePW=password	
	END IF
END FUNCTION

FUNCTION InsertRec(theSQL,theTable,theIDfield,theSelect)
	DIM theComm
    Set theComm = Server.CreateObject("ADODB.Command")
    theComm.ActiveConnection = MM_ecoinfo_STRING
    theComm.CommandText = theSQL
	theComm.Execute
    theComm.ActiveConnection.Close
	IF LEN(theTable)>0 THEN
		Set theComm = Server.CreateObject("ADODB.RecordSet")
		theComm.ActiveConnection = MM_ecoinfo_STRING
		theComm.Source = "SELECT MAX(" & theIDfield & ") as themax FROM " & theTable
		IF LEN(theSelect)>0 THEN
			theComm.Source = theComm.Source & " WHERE " & theSelect
		END IF
		theComm.Open()
		InsertRec=theComm.Fields.Item("themax").value
		theComm.Close()
	ELSE
		InsertRec=0
	END IF
END FUNCTION

SUB execCommand(theSQL)

	
	DIM theComm
    Set theComm = Server.CreateObject("ADODB.Command")
    theComm.ActiveConnection = MM_ecoinfo_STRING
    theComm.CommandText = theSQL
	
	theComm.Execute
    theComm.ActiveConnection.Close
END SUB

SUB MakeFile(theSQL,MM_ecoinfo_STRING,StartText,EndText,ModelText,Filename)
	DIM theRS, Format_Sel, fso, ts,theData,rowCount, colCount,i,ii, thisRow, theFile
	SET theRS= Server.CreateObject("ADODB.Recordset")
	theRS.ActiveConnection = MM_ecoinfo_STRING
	theRS.Source = theSQL
	theRS.Open()
	theData=theRS.GetRows
	theRS.Close()
	theRS=""
	colCount=uBound(theData,2)
	rowCount=uBound(theData,1)
	theFile=StartText
	FOR i=0 TO colCount
		thisRow=ModelText
		FOR ii=0 TO rowCount
			thisRow=replace(thisRow,"#" & ii & "#",trim(theData(ii,i)) & "")
		NEXT
		theFile=theFile & thisrow
	NEXT
	theFile=theFile & EndText
	WriteFile theFile,Filename
END SUB
Function FormatDate(theDate)
	IF LEN(theDate)>0 THEN
		theDate=CDate(theDate)
		FormatDate= DatePart("d",theDate) & "/" & DatePart("m",theDate) & "/" & DatePart("yyyy",theDate)
ELSE
	FormatDate="''"
	END IF
End Function

FUNCTION FormatDateDB(theDate)
	IF LEN(theDate)>0 THEN
		theDate=CDate(theDate)
		FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & datepart("m",theDate),2) & "-" &_
			right("0" & datepart("d",theDate),2) & " 00:00:00'}"
	ELSE
		FormatDateDB="''"
	END IF
END FUNCTION

FUNCTION SaveFile_ASP(UploadDirectory,newname)
	storeType="file"
	IF newname=1 THEN 
		nameConflict="uniq" 
	ELSE 
		nameConflict="over"
	END IF
	sizeLimit=""
	
	'Get the boundary
	PosBeg = 1
	PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))

' NAMES HERE

	GP_keys = UploadRequest.Keys
	for GP_i = 0 to UploadRequest.Count - 1
		GP_curKey = GP_keys(GP_i)
		'Save all uploaded files
		if UploadRequest.Item(GP_curKey).Item("FileName") <> "" AND UploadRequest.Item(GP_curKey).Item("ValueLen")>0 then
			GP_value = UploadRequest.Item(GP_curKey).Item("Value")
			GP_valueBeg = UploadRequest.Item(GP_curKey).Item("ValueBeg")
			GP_valueLen = UploadRequest.Item(GP_curKey).Item("ValueLen")
	
			GP_curPath = UploadDirectory
			GP_FullPath = Trim(Server.mappath(GP_curPath))

			'Create a Stream instance
			Dim GP_strm1, GP_strm2
			Set GP_strm1 = Server.CreateObject("ADODB.Stream")
			Set GP_strm2 = Server.CreateObject("ADODB.Stream")

			'Open the stream
			GP_strm1.Open
			GP_strm1.Type = 1 'Binary
			GP_strm2.Open
			GP_strm2.Type = 1 'Binary

			GP_strm1.Write RequestBin
			GP_strm1.Position = GP_ValueBeg
			GP_strm1.CopyTo GP_strm2,GP_ValueLen

			'Create and Write to a File
			GP_CurFileName = UploadRequest.Item(GP_curKey).Item("FileName")      
			GP_FullFileName = GP_FullPath & "\" & GP_CurFileName
			Set fso = CreateObject("Scripting.FileSystemObject")

			'Check if the file already exist
			GP_FileExist = false
			If fso.FileExists(GP_FullFileName) Then
				GP_FileExist = true
			End If      
			if ((nameConflict = "over" or nameConflict = "uniq") and GP_FileExist) or (NOT GP_FileExist) then
				if nameConflict = "uniq" and GP_FileExist then
					Begin_Name_Num = 0
					while GP_FileExist    
						Begin_Name_Num = Begin_Name_Num + 1
						GP_FullFileName = Trim(GP_FullPath)& "\" & fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
						GP_FileExist = fso.FileExists(GP_FullFileName)
					wend  
					UploadRequest.Item(GP_curKey).Item("FileName") = fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
					UploadRequest.Item(GP_curKey).Item("Value") = UploadRequest.Item(GP_curKey).Item("FileName")
				end if
				on error resume next
				GP_strm2.SaveToFile GP_FullFileName,2
				if err then
					Response.Write "<b>An error has occured saving uploaded file!</b><br><br>"
					Response.Write "Filename: " & Trim(GP_curPath) & UploadRequest.Item(GP_curKey).Item("FileName") & "<br>"
					Response.Write "Maybe the destination directory does not exist, or you don't have write permission.<br>"
					Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"
					err.clear
					GP_strm1.Close
					GP_strm2.Close
					response.End
				end if
				GP_strm1.Close
				GP_strm2.Close
				on error goto 0
			end if
			IF fileNames="" THEN
				filenames=UploadRequest.Item(GP_curKey).Item("FileName")
			ELSE
				filenames=fileNames & "," & UploadRequest.Item(GP_curKey).Item("FileName")			
			END IF
		end if
	next
	SaveFile_ASP=fileNames
End FUNCTION

'String to byte string conversion
Function getByteString(StringStr)
	For i = 1 to Len(StringStr)
		char = Mid(StringStr,i,1)
		getByteString = getByteString & chrB(AscB(char))
	Next
End Function

'Byte string to string conversion
Function getString(StringBin)
	getString =""
	For intCount = 1 to LenB(StringBin)
		getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
	Next
End Function

Function UploadFormRequest(name)
	on error resume next
	if UploadRequest.Item(name) then
		UploadFormRequest = UploadRequest.Item(name).Item("Value")
	end if  
End Function

Sub PureUploadSetup()
	UploadQueryString = Replace(Request.QueryString,"GP_upload=true","")
	if mid(UploadQueryString,1,1) = "&" then
		UploadQueryString = Mid(UploadQueryString,2)
	end if
	GP_uploadAction = CStr(Request.ServerVariables("URL")) & "?GP_upload=true"
	If (Request.QueryString <> "") Then  
		if UploadQueryString <> "" then
			GP_uploadAction = GP_uploadAction & "&" & UploadQueryString
		end if 
	End If

	'Get all data inside the boundaries
	PosBeg = 1
	PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
	boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
	boundaryPos = InstrB(1,RequestBin,boundary)
	i=1
	allNames=","
	Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))

		'Members variable of objects are put in a dictionary object
		Dim UploadControl
		Set UploadControl = CreateObject("Scripting.Dictionary")

		'Get an object name
		Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
		Pos = InstrB(Pos,RequestBin,getByteString("name="))
		PosBeg = Pos+6
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
		Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		WHILE (INSTR(allNames,"," & Name & ",")>0)
			Name=Name & CStr(i)
			i=i+1
		WEND
		allNames=allNames & Name & ","
		PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
		PosBound = InstrB(PosEnd,RequestBin,boundary)

		'Test if object is of file type
		If  PosFile<>0 AND (PosFile<PosBound) Then
			'Get Filename, content-type and content of file
			PosBeg = PosFile + 10
			PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			FileName = Mid(FileName,InStrRev(FileName,"\")+1)
			'Add filename to dictionary object
			UploadControl.Add "FileName", FileName
			Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
			PosBeg = Pos+14
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
			'Add content-type to dictionary object
			ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			UploadControl.Add "ContentType",ContentType
			'Get content of object
			PosBeg = PosEnd+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = FileName
			ValueBeg = PosBeg-1
			ValueLen = PosEnd-Posbeg
		Else
			'Get content of object
			Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
			PosBeg = Pos+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			ValueBeg = 0
			ValueEnd = 0
		End If

		'Add content to dictionary object
		UploadControl.Add "Value" , Value	
		UploadControl.Add "ValueBeg" , ValueBeg
		UploadControl.Add "ValueLen" , ValueLen	

		'Add dictionary object to main dictionary
		UploadRequest.Add name, UploadControl

		'Loop to next object
		BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
	Loop
End Sub

Function splitPhone(number)
	tlf=""
	nr=number
	length=Len(number)
	tal=""
	for i=1 to Len(number)
		a=Left(nr,1)
		length=length - 1
		nr=CStr(Right(nr,length))
		if IsNumeric(a) then
			tlf=tlf + CStr(a)
		end if
	next
	result=""
	nr=tlf
	length=Len(tlf)
	for i=1 to Len(tlf)
		result=result & Left(nr,1)
		length=length - 1
		nr=CStr(Right(nr,length))
		if i mod 2 = 0 then
			result=result & " "
		end if
	next
	splitPhone=result
End Function
Function toHTML(str)'fra input til SQL streng; tegnene: ' og " kodes
	s1=replace(str,Chr(34),"&quot;")
	s2=replace(s1,"'","&fnyt;")
	toHTML=s2
End Function
Function toString(str)'Fra Recordset til tekst; tegnet: ' 
	s1=replace(str,"&fnyt;","'")
	s2=replace(s1,"&quot;",Chr(34))
	s2=replace(s2,chr(13),"<br>")
	toString=s2
End Function

SUB SendCDOHtmlMail(theto,thefrom,theSubj,theBody)
	Set objMessage = CreateObject("CDO.Message") 
	objMessage.Subject = theSubj
	objMessage.From = thefrom 
	objMessage.To = theto
	objMessage.HTMLBody = theBody
	objMessage.Send
	set objMessage=nothing
END SUB
%>