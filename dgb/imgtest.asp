
<%
img="http://lh3.google.com/econetimages/R7JTCVSfHgI/AAAAAAAAAB4/OganHxSte4o/apples.jpg?"

function picasaImgSize(imgsrc,w)
i=1
imgname=Right(imgsrc,i) 
while InStr(imgname,"/")=0 
i=i+1
imgname=Right(imgsrc,i) 
wend
imgsrc=Left(imgsrc,Len(imgsrc)-i)
imgsrc=imgsrc & w & imgname
picasaImgSize=imgsrc
end function
%>
<img src="<%=picasaImgSize(img,"/s400")%>" width="360" />
