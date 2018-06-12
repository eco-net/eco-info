<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim wBranding,hBranding
wBranding=256
hBranding=200
Set Upload = Server.CreateObject("Persits.Upload.1")

' Upload files
Upload.OverwriteFiles = False ' Generate unique names
Upload.SetMaxSize 100000,true ' Truncate files above 1MB
Count=Upload.SaveVirtual("upload/") 

Set File = Upload.Files("file")

	if not File.ImageWidth=wBranding and not File.ImageHeight=hBranding then
		response.write("Branding skal være 258 * 200 pixel, men er: ")
		
		response.write File.ImageWidth & " * "
		response.write File.ImageHeight
		response.end
	end if

If Not File Is Nothing Then

		Set rs = Server.CreateObject("adodb.recordset")
		rs.Open "image_maindata", MM_ecoinfo_STRING, 2, 3
		
		rs.AddNew
		rs("branding") = File.Binary
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

