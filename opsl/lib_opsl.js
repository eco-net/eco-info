//////////////////////////////////////////////////////////////////////////////////////
//GENERAL UTILITIES

function checkImg(theField,theHeight,theWidth,ismax) {
	var error=0;
	var re = new RegExp(".(gif|jpg|png|bmp|jpeg)$","i");
	if(!re.test(theField.value) && theField.value.length>0){error=1;alert('The chosen image can not be displayed in a browser.')}
	if(error==0){
		var imgURL = 'file:///' + theField.value.replace(/\\/gi,'/');
		theField.gp_img = new Image();
		theField.gp_img.src=imgURL;
		if(!ismax){
			if(theField.gp_img.width!=theWidth && theWidth>0) error=1;
			if(theField.gp_img.height!=theHeight && theHeight>0) error=1;
		} else {
			if(theField.gp_img.width>theWidth && theWidth>0) error=1;
			if(theField.gp_img.height>theHeight && theHeight>0) error=1;			
		}
		if(error==1)alert('The chosen image is not of the correct dimensions.');
	};
	if(error==0)return true
	else {return false}
}

function saveImageDim(theField,theHeightField,theWidthField) {
		var imgURL = 'file:///' + theField.value.replace(/\\/gi,'/');
		theField.gp_img = new Image();
		theField.gp_img.src=imgURL;
		theWidthField.value=theField.gp_img.width
		theHeightField.value=theField.gp_img.height
}

function record(theform) {
	document.validation=false
	var theerr=0
	var imgOK
	if(theform.file.value.length>0){imgOK=checkImg(theform.file,500,500,true);saveImageDim(theform.file,theform.pheight,theform.pwidth)}
	else imgOK=true;
	if(imgOK==true){
		if(theform.titel.value.length==0 || theform.kort_beskrivelse.value.length==0 || theform.indsender.value.length==0  || theform.email.value.length==0 || theform.cat_id.selectedIndex==0) alert('Alle felter markeret med en * skal udfyldes.')
		else {document.validation=true}
	}
}

function donew(theform) {
	record(theform)
	if(document.validation==true) {
		theform.action='act_new.asp';theform.submit();
	}
}

