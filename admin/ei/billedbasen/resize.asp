<!-- AspUpload Code samples: resize.asp -->
<!-- Resizing JPEG images -->
<!-- Copyright (c) 2001 Persits Software, Inc. -->
<!-- http://www.persits.com -->

<HTML>
<BODY>

<h3>Resizing JPEG images</h3>
<h4>AspJpeg is required for this code sample</h4>

	<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="resize_upload.asp"> 
		Select a JPEG image:<BR>
		<INPUT TYPE="FILE" SIZE="40" NAME="FILE1"><BR>
		
		Scale: <SELECT NAME="size">
		<OPTION VALUE="branding">Branding</OPTION>
		<OPTION VALUE="splash">Splash</OPTION>
		<OPTION VALUE="page">Page</OPTION>
		</SELECT><P>
	<INPUT TYPE=SUBMIT VALUE="Upload!">
	</FORM>

</BODY>
</HTML>
