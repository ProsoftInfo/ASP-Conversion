function CheckSubmit() {
	gn = document.forms[0].txtClassName.value;
	/*
	if (document.forms[0].elements.length == 8) {
		if (document.forms[0].selCategory[document.forms[0].selCategory.selectedIndex].value == "select") {
			alert("Select Category");
			document.forms[0].selCategory.focus();
			return false;
		}
		else if (trimTrue(gn) == "") {
			alert("Enter Classification Name");
			document.forms[0].txtClassName.select();
			return false;
		}
		else {
			//return true;
			//document.forms[0].hItmType.value = document.forms[0].selItemType.value
			document.forms[0].action = "MasClassificationNameInsert.asp"
			document.forms[0].submit();
		}
	}
	else if(document.forms[0].elements.length == 6) {
		if (trimTrue(gn) == "") {
			alert("Enter Classification Name");
			document.forms[0].txtClassName.select();
			return false;
		}
		else {
			//document.forms[0].hItmType.value = document.forms[0].selItemType.value
			document.forms[0].action = "MasClassificationNameInsert.asp"
			document.forms[0].submit();
		}
	}
	*/
	if (document.forms[0].elements.length == 8) {
		if (document.forms[0].selCategory[document.forms[0].selCategory.selectedIndex].value == "select") {
			alert("Select Category");
			document.forms[0].selCategory.focus();
			return false;
		}
		else if (trimTrue(gn) == "") {
			alert("Enter Classification Name");
			document.forms[0].txtClassName.select();
			return false;
		}
		else {
			//return true;
			//document.forms[0].hItmType.value = document.forms[0].selItemType.value
			document.forms[0].action = "MasClassificationNameInsert.asp"
			document.forms[0].submit();
		}
	}
	else if(document.forms[0].elements.length == 6) {
		if (trimTrue(gn) == "") {
			alert("Enter Classification Name");
			document.forms[0].txtClassName.select();
			return false;
		}
		else {
			//document.forms[0].hItmType.value = document.forms[0].selItemType.value
			document.forms[0].action = "MasClassificationNameInsert.asp"
			document.forms[0].submit();
		}
	}
}

function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}
