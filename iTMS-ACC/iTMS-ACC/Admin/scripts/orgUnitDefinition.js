function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function checkSubmit(todaysdate){
	if (document.forms[0].selOrgUnit[document.forms[0].selOrgUnit.selectedIndex].value == "select") {
		alert("Select Organization Unit");
		document.forms[0].selOrgUnit.focus();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitName.value) == "") {
		alert("Enter Unit Name");
		document.forms[0].txtUnitName.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitShName.value) == "") {
		alert("Enter Organization Unit Short Name");
		document.forms[0].txtUnitShName.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitAddr1.value) == "") {
		alert("Enter Address");
		document.forms[0].txtUnitAddr1.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitPIN.value) == "") {
		alert("Enter PIN");
		document.forms[0].txtUnitPIN.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitCity.value) == "") {
		alert("Enter City");
		document.forms[0].txtUnitCity.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitState.value) == "") {
		alert("Enter State");
		document.forms[0].txtUnitState.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitPhone.value) == "") {
		alert("Enter Phone Number");
		document.forms[0].txtUnitPhone.select();
		return false;
	}
/*	else if (trimTrue(document.forms[0].txtUnitFax.value) == "") {
		alert("Enter Fax Number");
		document.forms[0].txtUnitFax.select();
		return false;
	}
*/
	else if (!checkmailid(document.forms[0].txtUnitEmail.value)) {
		return false;
	}
/*	
	else if (trimTrue(document.forms[0].txtUnitURL.value) == "") {
		alert("Enter Website URL");
		document.forms[0].txtUnitURL.select();
		return false;
	}
*/
	else if (trimTrue(document.forms[0].txtUnitContactPerson.value) == "") {
		alert("Enter Contact Person");
		document.forms[0].txtUnitContactPerson.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitTNGSTNo.value) == "") {
		alert("Enter TNGSTRC Number");
		document.forms[0].txtUnitTNGSTNo.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitAreaCode.value) == "") {
		alert("Enter Area Code");
		document.forms[0].txtUnitAreaCode.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitCSTRCNo.value) == "") {
		alert("Enter CSTRC Number");
		document.forms[0].txtUnitCSTRCNo.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitRange.value) == "") {
		alert("Enter Range");
		document.forms[0].txtUnitRange.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitDivision.value) == "") {
		alert("Enter Division");
		document.forms[0].txtUnitDivision.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitCollectorate.value) == "") {
		alert("Enter Collectorate");
		document.forms[0].txtUnitCollectorate.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitCentralENo.value) == "") {
		alert("Enter Central Excise Number");
		document.forms[0].txtUnitCentralENo.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitRegNo.value) == "") {
		alert("Enter Registration Number");
		document.forms[0].txtUnitRegNo.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtUnitLANo.value) == "") {
		alert("Enter LA Number");
		document.forms[0].txtUnitLANo.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtRangeAdd1.value) == "") {
		alert("Enter Range Address");
		document.forms[0].txtRangeAdd1.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtDivisionAdd1.value) == "") {
		alert("Enter Division Address");
		document.forms[0].txtDivisionAdd1.select();
		return false;
	}
	else {
		document.forms[0].action = "orgUnitDefinitionInsert.asp"
		document.forms[0].hcountry.value = document.forms[0].selUnitCountry.options[document.forms[0].selUnitCountry.selectedIndex].text;
		sethiddenDate()
		document.forms[0].submit();
	}
}

function checkmailid(mailid) {
	var exclude=/[^@\-\.\w]{2}|[@\.]{2}|(@)[^@]*\1/;
	var check=/@[\w\-]+\./;
	var checkend=/\.[a-zA-Z]{2,3}$/;

	if (trimTrue(mailid) == "") {
		alert("Enter Emailid");
		document.forms[0].txtUnitEmail.select();
		return false;
	}
	else if(((mailid.search(exclude) != -1)||(mailid.search(check))==-1)||(mailid.search(checkend) == -1)) {
		alert("Enter Valid Emailid");
		document.forms[0].txtUnitEmail.select();
		return false;
	}
	else
		return true;
}
