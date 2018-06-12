<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim wBranding,hBranding,wSplash,hSplash,w2col,w3col,wRigthcol
w2col=540
w3col=360
wRightcol=160
wSplash=128
hSplash=100
wBranding=256
hBranding=200

Set Upload = Server.CreateObject("Persits.Upload.1")

' Upload files
Upload.OverwriteFiles = False ' Generate unique names
Upload.SetMaxSize 100000,true ' Truncate files above 1MB
Count=Upload.SaveVirtual("upload/") 
Set File = Upload.Files("file")
if Upload.Form("size")= "1" then
	if not File.ImageWidth=wBranding and not File.ImageHeight=hBranding then
		response.write("Branding skal være " & wBranding & " * " & hBranding & " pixel, men er: ")
		response.write File.ImageWidth & " * "
		response.write File.ImageHeight
		response.end
	end if
end if
if Upload.Form("size")= "2" then
	if not File.ImageWidth=wSplash and not File.ImageHeight=hSplash then
		response.write("Splash skal være " & wSplash & " * " & hSplash & " pixel, men er: ")
		response.write File.ImageWidth & " * "
		response.write File.ImageHeight
		response.end
	end if
end if
if Upload.Form("size")= "4" then 'Rightcol

	if  File.ImageWidth > wRightcol then
		response.write "Rightcol må højst være " & wRightcol & " pixel, men er: "
		response.write File.ImageWidth 
		response.end
	end if
end if
if Upload.Form("size")= "5" then '3col
	if  File.ImageWidth > w3col then
		response.write("Main3col må højst være " & w3col & " pixel, men er: ")
		response.write File.ImageWidth 
		response.end
	end if
end if
if Upload.Form("size")= "6" then '2col
	if  File.ImageWidth > w2col then
		response.write("Main2col må højst være " & w2col & " pixel, men er: ")
		response.write File.ImageWidth 
		response.end
	end if
end if
If Not File Is Nothing Then
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Open "image_maindata", MM_ecoinfo_STRING, 2, 3
		rs.AddNew
		rs("image") = File.Binary
		rs("size_id") = CInt(Upload.Form("size"))
		rs("cat_id") = CInt(Upload.Form("cat"))
		rs("filename") = File.ExtractFileName
		rs("kb") = File.Size/1024
		rs("subtext") = Upload.Form("subtext")
		rs("source") = Upload.Form("source")
		rs("ext")= File.ImageType
		rs("height")= File.ImageHeight
		rs("width")= File.ImageWidth
		rs.Update
		rs.Close
		rs.Open "image_maindata", MM_ecoinfo_STRING, 2, 3
		rs.MoveLast
		theID=rs("id")
		rs.Close
		Response.Write "Billedet er gemt i databasen, som nr: " & theID & ".<br>"
		Response.Write File.ExtractFileName & "<br>" 
		old=File.ExtractFileName
		path= "original/" & CStr(theID) & "." & File.ImageType
		response.write("kopi gemmes som: " & path & "<br>")
		File.CopyVirtual(path)
		Response.Write "<img src='" & path & "'>"
	Else
		Response.Write "File not selected."
End If


  File.Delete  
  

%> 

