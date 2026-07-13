function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function checkSubmit(){
	if (trimTrue(document.forms[0].txtUOMCode.value) == "") {
		alert("Enter UoM Short Description");
		document.forms[0].txtUOMCode.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUOMName.value) == "") {
		alert("Enter UoM Description");
		document.forms[0].txtUOMName.select();
		return false;
	}
	else if (!(document.forms[0].radDecimal[0].checked || document.forms[0].radDecimal[1].checked)) {
		alert("Select Decimals allowed");
		document.forms[0].radDecimal[0].focus();
		return false;
	}
	else {
		document.forms[0].action = "masUOMInsert.asp"
		document.forms[0].submit();
	}
}
