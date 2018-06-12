<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="lib.asp" -->

<%



'**** Saving imagefile
'DIM theImagePath, UploadRequest,filename
'Set UploadRequest = CreateObject("Scripting.Dictionary")
'if request("file")<>"" then
'DIM  RequestBin
'RequestBin = Request.BinaryRead(Request.TotalBytes)
'PureUploadSetup
'theImagePath="/log/ei/opsl/graphics/"
'filename=theImagePath & SaveFile_ASP(theImagePath,1)
'else
'filename=""
'end if
DoSave filename,1

%>
