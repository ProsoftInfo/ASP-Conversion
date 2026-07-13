function CheckSubmit() {
	gn = document.forms[0].txtClassName.value;
	if (trimTrue(gn) == "") {
		alert("Enter Classification Name");
		document.forms[0].txtClassName.select();
		return false;
	}
	else {
		document.forms[0].action = "MasClassificationNameUpdate.asp"
		document.forms[0].submit();
	}
}

function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}
