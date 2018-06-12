<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim econet
econet="http://www.eco-net.dk/billedbasen/"
Dim wBranding,hBranding,wSplash,hSplash,w2col,w3col,wRigthcol,wTema
w2col=540
w3col=360
wRightcol=160
wSplash=128
hSplash=100
wBranding=256
hBranding=200
wTema=256
Dim branding,splash,threecol,twocol,tema,rightcol,path
branding=0
splash=0
threecol=0
twocol=0
tema=0
rightcol=0
path=""

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
	else
		branding=1
		path="branding/"
	end if
end if
if Upload.Form("size")= "2" then
	if not File.ImageWidth=wSplash and not File.ImageHeight=hSplash then
		response.write("Splash skal være " & wSplash & " * " & hSplash & " pixel, men er: ")
		response.write File.ImageWidth & " * "
		response.write File.ImageHeight
		response.end
	else
		splash=1
		path="splash/"
	end if
end if
if Upload.Form("size")= "4" then 'Rightcol

	if  File.ImageWidth > wRightcol then
		response.write "Rightcol må højst være " & wRightcol & " pixel, men er: "
		response.write File.ImageWidth 
		response.end
	else
		rightcol=1
		path="rightcol/"
	end if
end if
if Upload.Form("size")= "5" then '3col
	if  File.ImageWidth > w3col then
		response.write("Main3col må højst være " & w3col & " pixel, men er: ")
		response.write File.ImageWidth 
		response.end
	else
		threecol=1
		path="3col/"
	end if
end if
if Upload.Form("size")= "6" then '2col
	if  File.ImageWidth > w2col then
		response.write("Main2col må højst være " & w2col & " pixel, men er: ")
		response.write File.ImageWidth 
		response.end
	else
		twocol=1
		path="3col/"
	end if
end if
if Upload.Form("size")= "7" then 'tema
	if  File.ImageWidth > wTema then
		response.write("Tema må højst være " & wTema & " pixel, men er: ")
		response.write File.ImageWidth 
		response.end
	else
		tema=1
		path="tema/"
	end if
end if

If Not File Is Nothing Then
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Open "images", MM_ecoinfo_STRING, 2, 3
		rs.AddNew
		rs("filename") = File.ExtractFileName
		rs("subtext") = Upload.Form("subtext")
		rs("source") = Upload.Form("source")
		rs("cat_id") = Upload.Form("cat")
		rs("branding")= branding
		rs("splash")= splash
		rs("twocol")= twocol
		rs("threecol")= threecol
		rs("tema")= tema
		rs("rightcol")= rightcol
		rs("thumbnail")= 0
		rs.Update
		rs.Close
		rs.Open "images", MM_ecoinfo_STRING, 2, 3
		rs.MoveLast
		theID=rs("id")
		rs.Close
		Response.Write "Billedet er gemt i databasen, som nr: " & theID & ".<br>"
		Response.Write File.ExtractFileName & "<br>" 
		path=path & File.ExtractFileName
		response.write("kopi gemmes som: " & path & "<br>")
		File.CopyVirtual(path)
		Response.Write "<img src='" & path & "'>"
	Else
		Response.Write "File not selected."
End If

%> 

