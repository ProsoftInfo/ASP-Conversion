function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function checkSubmit(){
	if (trimTrue(document.forms[0].txtUnitName.value) == "") {
		alert("Enter Organization Unit Name");
		document.forms[0].txtUnitName.select();
		return false;
	}
	else {
		document.forms[0].action = "orgUnitCreationInsert.asp"
		document.forms[0].submit();
	}
}
