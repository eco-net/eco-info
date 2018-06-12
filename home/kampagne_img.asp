<!--#include virtual="/shared/showimage.asp" -->
<%
if request("img")="1" then
file="graphics/folder_side1.jpg"
end if
if request("img")="2" then
file="graphics/folder_side2.jpg"
end if
if request("img")="3" then
file="graphics/plakat.jpg"
end if
%>
<html>
<head>
<title>Folder</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<img src="<%=file%>"> 
</body>
</html>
