function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function checkNumbers(val){
	var valid = "0123456789"
	var temp;

	for (var i=0; i < val.length; i++) {
		temp = "" + val.substring(i, i+1);
		if (valid.indexOf(temp) == "-1") 
			return false;
		else
			return true;
	}
}

function checkSubmit(todaysdate){
	if (trimTrue(document.forms[0].txtOrgName.value) == "") {
		alert("Enter Organization Name");
		document.forms[0].txtOrgName.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgShortName.value) == "") {
		alert("Select Organization Short Name");
		document.forms[0].txtOrgShortName.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgNoUnits.value) == "") {
		alert("Enter Number of Units");
		document.forms[0].txtOrgNoUnits.select();
		return false;
	}
	else if (!checkNumbers(document.forms[0].txtOrgNoUnits.value)) {
		alert("Enter Numerals Only");
		document.forms[0].txtOrgNoUnits.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgAddress1.value) == "") {
		alert("Enter Address");
		document.forms[0].txtOrgAddress1.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgPIN.value) == "") {
		alert("Enter PIN");
		document.forms[0].txtOrgPIN.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgCity.value) == "") {
		alert("Enter City");
		document.forms[0].txtOrgCity.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgState.value) == "") {
		alert("Enter State");
		document.forms[0].txtOrgState.select();
		return false;
	}
	else if (document.forms[0].selOrgCountry[document.forms[0].selOrgCountry.selectedIndex].value == "select") {
		alert("Select Country");
		document.forms[0].selOrgCountry.focus();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgPhone.value) == "") {
		alert("Enter Phone Number");
		document.forms[0].txtOrgPhone.select();
		return false;
	}
/*	else if (trimTrue(document.forms[0].txtOrgFax.value) == "") {
		alert("Enter Fax Number");
		document.forms[0].txtOrgFax.select();
		return false;
	}
*/
	else if (!checkmailid(document.forms[0].txtOrgEmail.value)) {
		return false;
	}
/*	
	else if (trimTrue(document.forms[0].txtOrgURL.value) == "") {
		alert("Enter Website URL");
		document.forms[0].txtOrgURL.select();
		return false;
	}
*/
	else if (trimTrue(document.forms[0].txtOrgContactPerson.value) == "") {
		alert("Enter Contact Person");
		document.forms[0].txtOrgContactPerson.select();
		return false;
	}
	else if (document.forms[0].selOrgCurrency[document.forms[0].selOrgCurrency.selectedIndex].value == "select") {
		alert("Select Operating Currency");
		document.forms[0].selOrgCurrency.focus();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgTNGSTRCNo.value) == "") {
		alert("Enter TNGSTRC Number");
		document.forms[0].txtOrgTNGSTRCNo.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgAreaCode.value) == "") {
		alert("Enter Area Code");
		document.forms[0].txtOrgAreaCode.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtOrgCSTRCNo.value) == "") {
		alert("Enter CSTRC Number");
		document.forms[0].txtOrgCSTRCNo.select();
		return false;
	}
	else {
		document.forms[0].action = "orgCreationInsert.asp"
		document.forms[0].hcountry.value = document.forms[0].selOrgCountry.options[document.forms[0].selOrgCountry.selectedIndex].text;
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
		document.forms[0].txtOrgEmail.select();
		return false;
	}
	else if(((mailid.search(exclude) != -1)||(mailid.search(check))==-1)||(mailid.search(checkend) == -1)) {
		alert("Enter Valid Emailid");
		document.forms[0].txtOrgEmail.select();
		return false;
	}
	else
		return true;
}
