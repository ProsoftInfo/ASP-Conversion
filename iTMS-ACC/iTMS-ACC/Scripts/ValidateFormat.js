function DoKeyPress(sYesNo,iIntPart,iDecPart) {
	var sIntVal
	sIntVal=""
	eTD = window.event.srcElement;
	
	if (sYesNo == "N") {
		if (window.event.keyCode < 48 || window.event.keyCode > 57) {
			window.event.keyCode ="\b";
		}
	}
	else if (sYesNo == "Y") {
		if ((window.event.keyCode < 48 || window.event.keyCode > 57) && window.event.keyCode != 46) {
			window.event.keyCode ="\b";
		}
	}
	
	sValue = new String(eTD.value);
	
	iDecPostion = sValue.indexOf(".");
	
	if (iDecPostion >= 0) {
		sDecVal = sValue.substring(iDecPostion + 1,sValue.length);
		sIntVal = sValue.substring(0,iDecPostion);
	}
	else {
		sDecVal="";
		sIntVal = sValue
	}

	if (sYesNo == "N") {
		if (sIntVal.length >= iIntPart)
			window.event.keyCode = "\b";
	}
	else if (sYesNo == "Y") {
		if (iDecPostion >= 0) {
			if (window.event.keyCode == 46 || (sDecVal.length >= iDecPart))
				window.event.keyCode = "\b";
		}
		else {
			if (sIntVal.length = iIntPart) {
				if (sDecVal.length >= iDecPart)
					window.event.keyCode = "\b";
			}
			if ((sIntVal.length >= iIntPart) && window.event.keyCode != 46)
				window.event.keyCode = "\b";
			
		}
		
	}
}	

function DoKeyPressHypen(sYesNo,iIntPart,iDecPart) {
	var sIntVal
	sIntVal=""
	eTD = window.event.srcElement;
	
	if (sYesNo == "N") {
		if (window.event.keyCode < 48 || window.event.keyCode > 57 && window.event.keyCode != 45) {
			window.event.keyCode ="\b";
		}
	}
	else if (sYesNo == "Y") {
		if ((window.event.keyCode < 48 || window.event.keyCode > 57) && window.event.keyCode != 46 && window.event.keyCode != 45) {
			window.event.keyCode ="\b";
		}
	}
	
	sValue = new String(eTD.value);
	
	iDecPostion = sValue.indexOf(".");
	
	if (iDecPostion >= 0) {
		sDecVal = sValue.substring(iDecPostion + 1,sValue.length);
		sIntVal = sValue.substring(0,iDecPostion);
	}
	else {
		sDecVal="";
		sIntVal = sValue
	}

	if (sYesNo == "N") {
		if (sIntVal.length >= iIntPart)
			window.event.keyCode = "\b";
	}
	else if (sYesNo == "Y") {
		if (iDecPostion >= 0) {
			if (window.event.keyCode == 46 || (sDecVal.length >= iDecPart))
				window.event.keyCode = "\b";
		}
		else {
			if (sIntVal.length = iIntPart) {
				if (sDecVal.length >= iDecPart)
					window.event.keyCode = "\b";
			}
			if ((sIntVal.length >= iIntPart) && window.event.keyCode != 46)
				window.event.keyCode = "\b";
			
		}
		
	}
}	

function CheckAlpha(len) {
var flag, kcode, sLen, sStr;
kcode = window.event.keyCode;
sLen = eval(len);
sStr = new Array();
sStr = window.event.srcElement.value;
flag = true;
	if ( kcode >= 97 && kcode <= 122 )
	{
		flag = false;
	}
	else if ( kcode >= 65 && kcode <= 90 )
	{
		flag = false;
	}
	else if( flag == true )
	{
		window.event.keyCode ="\b";
		alert("Only Alphabets should be entered");
	}


	if ( sLen <= sStr.length )
	{
		window.event.keyCode ="\b";
	}

}

function DoKeyPressText(iLength) {
	var sTextVal
	sTextVal = ""
	sTextVal = window.event.srcElement.value;
	
	if ( iLength <= sTextVal.length )
	{
		window.event.srcElement.value = sTextVal.substring(0,iLength);
	}
}	