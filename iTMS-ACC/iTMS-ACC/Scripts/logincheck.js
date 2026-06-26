function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function validate(){
	var pass = document.forms[0].password.value;
	var passretype = document.forms[0].retypepassword.value;
	var checkCheck = false;
	if (!loginidcheck()) {
		return false;
	}
	else if (trimTrue(pass) == "") {
		alert("Enter Password");
		document.forms[0].password.select();
		return false;
	}
	else if (trimTrue(passretype) == "") {
		alert("Re-Type Password");
		document.forms[0].retypepassword.select();
		return false;
	}
	else if (pass.length < 8) {
		alert("Password should be aleast 8 characters");
		document.forms[0].password.select();
		return false;
	}
	else if (trimTrue(pass) != trimTrue(passretype)) {
		alert("Password and Re-Entered Password should be same");
		document.forms[0].password.select();
		return false;
	}
	else if (trimTrue(document.forms[0].Username.value) == "") {
		alert("Enter Name");
		document.forms[0].Username.select();
		return false;
	}
	else if (trimTrue(document.forms[0].Qualification.value) == "") {
		alert("Enter Qualification");
		document.forms[0].Qualification.select();
		return false;
	}
	else if (trimTrue(document.forms[0].Clinicalattachment.value) == "") {
		alert("Enter Clinical Attachment");
		document.forms[0].Clinicalattachment.select();
		return false;
	}
	else if (trimTrue(document.forms[0].Postaladdress.value) == "") {
		alert("Enter Postal Address");
		document.forms[0].Postaladdress.focus();
		return false;
	}
	else if (trimTrue(document.forms[0].city.value) == "") {
		alert("Enter City");
		document.forms[0].city.select();
		return false;
	}
	else if (trimTrue(document.forms[0].state.value) == "") {
		alert("Enter State");
		document.forms[0].state.select();
		return false;
	}
/*	else if (trimTrue(document.forms[0].country.value) == "") {
		alert("Enter Country");
		document.forms[0].country.select();
		return false;
	}*/
	else if (trimTrue(document.forms[0].postalcode.value) == "") {
		alert("Enter PIN Code");
		document.forms[0].postalcode.select();
		return false;
	}
	else if (isNaN(document.forms[0].postalcode.value)) {
		alert("Invalid PIN Code");
		document.forms[0].postalcode.select();
		return false;
	}
	else if (trimTrue(document.forms[0].phonecode.value) == "") {
		alert("Enter Phone Number");
		document.forms[0].phonecode.select();
		return false;
	}
	else if (!phnocheck()) {
		return false;
	}
	else if (!checkmailid(document.forms[0].email.value)) {
		return false;
	}
	else if ((!(document.forms[0].Fieldsofimaging[0].checked)) && (!(document.forms[0].Fieldsofimaging[1].checked)) && (!(document.forms[0].Fieldsofimaging[2].checked)) && (!(document.forms[0].Fieldsofimaging[3].checked)) && (!(document.forms[0].Fieldsofimaging[4].checked)) && (!(document.forms[0].Fieldsofimaging[5].checked))) {
		alert("Check Fields of Imaging");
		return false;
	}
	else if ((!(document.forms[0].areasofinterest[0].checked)) && (!(document.forms[0].areasofinterest[1].checked)) && (!(document.forms[0].areasofinterest[2].checked)) && (!(document.forms[0].areasofinterest[3].checked)) && (!(document.forms[0].areasofinterestOthers.checked))) {
		alert("Check Areas of Interest");
		return false;
	}
/*	else {
		for (i = 0;i<6;i++) {
			foicheckvalue = document.forms[0].Fieldsofimaging[i].checked;
			if (foicheckvalue) {
				foicheckvalue = document.forms[0].Fieldsofimaging[i].value;
				document.forms[0].Fieldsofimagingh.value = foicheckvalue
				//break;
			}
		}

		for (j = 0;j<4;j++) {
			aoicheckvalue = document.forms[0].areasofinterest[j].checked;
			if (aoicheckvalue) {
				checkCheck  =true;
				aoicheckvalue = document.forms[0].areasofinterest[j].value;
				}
				document.forms[0].areasofinteresth.value = aoicheckvalue
				//break;
			}
		}
		
		if (!checkCheck) {
			if (trimTrue(document.forms[0].other.value) == "") {
				alert("Enter Areas of Interest");
				document.forms[0].other.select();
				return false;
				//break;
			}
			else
				aoicheckvalue = document.forms[0].other.value;	
		return true;
	}
*/	

	if (document.forms[0].areasofinterestOthers.checked && trimTrue(document.forms[0].other.value) == "") {
		alert("Enter Areas of Interest");
		document.forms[0].other.select();
		return false;
	}
	return true;


}

function loginidcheck() {
	var name;
	name = document.forms[0].Loginid.value;
	var loginvalid = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._";
	var loginok = "yes";
	var temp_login;
	for (var i = 0; i < name.length; i++) {
		temp_login = "" + name.substring(i, i+1);
		if (loginvalid.indexOf(temp_login) == "-1")
			loginok = "no";
	}
	if (trimTrue(document.forms[0].Loginid.value) == "") {
		alert("Enter Loginid");
		document.forms[0].Loginid.select();
		return false;
	}
	else if (loginok == "no") {
		alert("Enter characters from a-z or 0-9 or . or _ for Loginid");
		document.forms[0].Loginid.select();
		return false;
	}
	else
		return true;
}

function checkmailid(mailid) {
	var exclude=/[^@\-\.\w]{2}|[@\.]{2}|(@)[^@]*\1/;
	var check=/@[\w\-]+\./;
	var checkend=/\.[a-zA-Z]{2,3}$/;

	if (trimTrue(mailid) == "") {
		alert("Enter Emailid");
		document.forms[0].email.select();
		return false;
	}
	else if(((mailid.search(exclude) != -1)||(mailid.search(check))==-1)||(mailid.search(checkend) == -1)) {
		alert("Enter Valid Emailid");
		document.forms[0].email.select();
		return false;
	}
	else
		return true;
}

function passcheck() {
	if (trimTrue(document.forms[0].oldpass.value) == "") {
		alert("Enter Old Password");
		document.forms[0].oldpass.select();
		return false;
	}
	else if (trimTrue(document.forms[0].newpass.value) == "") {
		alert("Enter New Password");
		document.forms[0].newpass.select();
		return false;
	}
	else if (trimTrue(document.forms[0].conpass.value) == "") {
		alert("Confirm Password");
		document.forms[0].conpass.select();
		return false;
	}
	else if (trimTrue(document.forms[0].newpass.value) != trimTrue(document.forms[0].conpass.value)) {
		alert("Password and Confirm Password should be same");
		document.forms[0].newpass.select();
		return false;
	}
	else
		return true;
}

function phnocheck() {
	var name;
	name = document.forms[0].phonecode.value;
	
	var loginvalid = "1234567890-,";
	var loginok = "yes";
	var temp_login;
	for (var i = 0; i < name.length; i++) {
		temp_login = "" + name.substring(i, i+1);
		if (loginvalid.indexOf(temp_login) == "-1")
			loginok = "no";
	}

	if (loginok == "no") {
		alert("Invalid Phone Number, only numbers and hyphen(-) are allowed.");
		document.forms[0].phonecode.select();
		return false;
	}
	else
		return true;
}
