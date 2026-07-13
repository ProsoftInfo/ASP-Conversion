function CheckSubmit() {
	aname = document.forms[0].txtAttrName.value;

	if (trimTrue(aname) == "")
	{
		alert("Enter Attribute Name");
		document.forms[0].txtAttrName.select();
		return false;
	}
	else if (document.forms[0].selDataType[document.forms[0].selDataType.selectedIndex].value == "select") {
		alert("Select Data Type");
		document.forms[0].selDataType.focus();
		return false;
	}
	else if (!datalenCheck()) {
		;
	}
	else if (!document.forms[0].txtDecimal.disabled && !decilenCheck()) {
		;
	}
	else {
		document.forms[0].action = "MasClassificationAttributeInsert.asp"			
		document.forms[0].submit();
	}
}

function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function datalenCheck() {
	var name;
	name = document.forms[0].txtDataLen.value;

	var valid = "1234567890";
	var ok = "yes";
	var temp;
	for (var i = 0; i < name.length; i++) {
		temp = "" + name.substring(i, i+1);
		if (valid.indexOf(temp) == "-1")
			ok = "no";
	}

	if (trimTrue(document.forms[0].txtDataLen.value) == "")
	{
		alert("Enter Attribute Data Length");
		document.forms[0].txtDataLen.select();
		return false;
	}
	else if (ok == "no") {
		alert("Only numerals are allowed.");
		document.forms[0].txtDataLen.select();
		return false;
	}
	else
		return true;
}

function decilenCheck() {
	var name;
	name = document.forms[0].txtDecimal.value;

	var valid = "1234567890";
	var ok = "yes";
	var temp;
	for (var i = 0; i < name.length; i++) {
		temp = "" + name.substring(i, i+1);
		if (valid.indexOf(temp) == "-1")
			ok = "no";
	}

	if (ok == "no") {
		alert("Only numerals are allowed.");
		document.forms[0].txtDecimal.select();
		return false;
	}
	else
		return true;
}

function checkSelect() {
	if (document.forms[0].selDataType[document.forms[0].selDataType.selectedIndex].value == "String") {
		document.forms[0].txtDataLen.disabled = false;
		document.forms[0].txtDecimal.disabled = true;
	}
	else if (document.forms[0].selDataType[document.forms[0].selDataType.selectedIndex].value == "Numeric") {
		document.forms[0].txtDataLen.disabled = false;
		document.forms[0].txtDecimal.disabled = false;
	}
	else if (document.forms[0].selDataType[document.forms[0].selDataType.selectedIndex].value == "Date Time") {
		document.forms[0].txtDataLen.disabled = false;
		document.forms[0].txtDecimal.disabled = true;
	}
	else if (document.forms[0].selDataType[document.forms[0].selDataType.selectedIndex].value == "Boolean") {
		document.forms[0].txtDataLen.disabled = false;
		document.forms[0].txtDecimal.disabled = true;
	}
}
