<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim wSplash,hSplash
wSplash=128
hSplash=100
Set Upload = Server.CreateObject("Persits.Upload.1")

' Upload files
Upload.OverwriteFiles = False ' Generate unique names
Upload.SetMaxSize 100000,true ' Truncate files above 1MB
Count=Upload.SaveVirtual("upload/") 

Set File = Upload.Files("file")

	if not File.ImageWidth=wSplash and not File.ImageHeight=hSplash then
		response.write("Splash skal være " & wSplash & " * " & hSplash & " pixel, men er: ")
		response.write File.ImageWidth & " * "
		response.write File.ImageHeight
		response.end
	end if

If Not File Is Nothing Then

		Set rs = Server.CreateObject("adodb.recordset")
		rs.Open "image_maindata", MM_ecoinfo_STRING, 2, 3
		
		rs.AddNew
		rs("splash") = File.Binary
		rs("cat_id") = CInt(Upload.Form("cat"))
		rs("filename") = File.ExtractFileName
		rs("kb") = File.Size/1024
		rs("subtext") = Upload.Form("subtext")
		rs("source") = Upload.Form("source")
		'rs("ext")= File.Ext 
		rs("height") = File.ImageHeight
		rs("width") = File.ImageWidth
		rs.Update
		
		Response.Write "File saved."
	Else
		Response.Write "File not selected."
End If
 	  

%> 

