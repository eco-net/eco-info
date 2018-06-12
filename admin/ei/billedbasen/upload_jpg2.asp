<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim w2col,w3col,wTema,wRightcol,wThumbnail,width,height,path,theID,filename
w2col=540
w3col=360
wTema=258
wRightcol=160
wThumbnail=75

Set Upload = Server.CreateObject("Persits.Upload.1")
' Upload files
Upload.OverwriteFiles = False ' Generate unique names
Upload.SetMaxSize 100000,true ' Truncate files above 1MB
Count=Upload.SaveVirtual("upload/") 
Set File = Upload.Files("file")
filename=File.ExtractFileName
response.write("OK")
response.end

if File.ext =".jpg" then

width = File.ImageWidth
height = File.ImageHeight
path=Server.MapPath("images/") & "/" & filename

Set Image = Server.CreateObject("AspImage.Image")
	if width >w3col then
		if width >=w2col then 
			makeImage(w2col,w2col) ' make size = w2col
		else 
			makeImage(width,w2col) 'make size max w2col 
		end if
	end if
	if width >wTema then 
		if width =w3col then 
			makeImage(w3col,w3col) ' make size = w3col
		else
			makeImage(width,w3col) ' make size max w3col
		end if
	end if
	if width >wRightcol then 
		if File.ImageWidth =wTema then 
			makeImage(wTema,wTema) ' make size = wTema
		else
			makeImage(width,wTema) ' make size max wTema
		end if
	end if
	if width >wThumbnail then 
		if File.ImageWidth =wRightcol then 
			makeImage(wRightcol,wRightcol) ' make size = wRightcol
		else
			makeImage(width,wRightcol) ' make size max wRightcol
		end if
	end if
	if width >=wThumbnail then 
			makeImage(wThumbnail,wThumbnail) ' make size = wThumbnail
	end if
else 'not .jpg
 response.write("Virker kun på .jpg filer<br>")
end if '.jpg

Function writeImageDescr()
If Not File Is Nothing Then
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Open "image_maindata", MM_ecoinfo_STRING, 2, 3
		rs.AddNew
'		rs("image") = File.Binary
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
	Else
		Response.Write "File not selected."
End If

end Function

Function makeImage(theSize,sizeName)
	response.write("theSize: " & theSize & "-sizeName: " & sizeName & "<br>")
	Image.LoadImage(path)
	x= theSize/width * height
	Image.Resize theSize,x
	Image.Filename=Server.MapPath("smallimages/")  & "/" & filename
	Image.SaveImage

End Function


'		Response.Write "Billedet er gemt i databasen, som nr: " & theID & ".<br>"
'		Response.Write File.ExtractFileName & "<br>" 
'		old=File.ExtractFileName
'		path= "original/" & CStr(theID) & "." & File.ImageType
'		response.write("kopi gemmes som: " & path & "<br>")
'		File.CopyVirtual(path)
'		Response.Write "<img src='" & path & "'>"
'	Else
'		Response.Write "File not selected."
'End If


 ' File.Delete  
  

%> 

