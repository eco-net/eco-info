<!-- AspUpload Code samples: resize_upload.asp -->
<!-- Invoked by resize.asp -->
<!-- Copyright (c) 2001 Persits Software, Inc. -->
<!-- http://www.persits.com -->

<HTML>
<BODY>

<%
Dim w2col,w3col,wRigthcol

w2col=530
w3col=360
wRightcol=160

Set Upload = Server.CreateObject("Persits.Upload.1")
' Use AspJpeg to resize image
Set Jpeg = Server.CreateObject("Persits.Jpeg")
' Capture and save uploaded image
Upload.Save Server.MapPath("upload/")
For Each File in Upload.Files
	If File.ImageType <> "JPG" Then
		Response.Write "This is not a JPEG image."
		File.Delete
		Response.End
	End If

	Jpeg.Open File.Path
	'size = Upload.Form("size")
	'if size="page" then
		if Jpeg.OriginalWidth>=w2col then
		response.write("w2col")
			Jpeg.Width = w2col
			Jpeg.Height = CInt(Jpeg.OriginalHeight * w2col / Jpeg.OriginalWidth)
			Jpeg.Save Server.MapPath("2col/" & File.ExtractFileName )
		end if
		if Jpeg.OriginalWidth>=w3col then
		response.write("w3col")
			Jpeg.Width = w3col
			Jpeg.Height = CInt(Jpeg.OriginalHeight * w3col / Jpeg.OriginalWidth)
			Jpeg.Save Server.MapPath("3col/" & File.ExtractFileName )
		end if
		if Jpeg.OriginalWidth>=wRigthcol then
		response.write("wRightcol")
			Jpeg.Width = wRightcol
			Jpeg.Height = CInt(Jpeg.OriginalHeight * wRightcol / Jpeg.OriginalWidth)
			Jpeg.Save Server.MapPath("rightcol/" & File.ExtractFileName )
		end if
	
	'end if
%>
	<IMG SRC="upload/<% = File.ExtractFileName %>"><BR>
	<IMG SRC="2col/<% = File.ExtractFileName %>"><br>
		<IMG SRC="3col/<% = File.ExtractFileName %>"><br>
		<IMG SRC="rightcol/<% = File.ExtractFileName %>"><br>
<%
Next
%>

</BODY>
</HTML>
