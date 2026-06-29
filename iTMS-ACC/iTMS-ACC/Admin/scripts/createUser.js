function CheckSubmit() {
	ctr = 0;

	if(trimTrue(document.forms[0].txtLoginID.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter Login ID");
		document.forms[0].txtLoginID.select();
	}
	else if (document.forms[0].selUnit.selectedIndex == 0 )
	{
	    ctr=ctr+1;
	    alert("Select Organization");
	    document.forms[0].selUnit.focus;
	}
	else if (document.forms[0].radUserType(0).checked)
	{
		document.formname.hUserType.value = document.forms[0].radUserType(0).value;
		
	}
	else if (document.forms[0].radUserType(1).checked)
	{
		document.formname.hUserType.value = document.forms[0].radUserType(1).value;
		
	}
	else if (document.forms[0].radUserType(2).checked)
	{
		document.formname.hUserType.value = document.forms[0].radUserType(2).value;
		
	}
	else if(trimTrue(document.forms[0].txtPassword.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter Password");
		document.forms[0].txtPassword.select();
	}
	else if(trimTrue(document.forms[0].txtFName.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter First Name");
		document.forms[0].txtFName.select();
	}
	else if(trimTrue(document.forms[0].txtLName.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter Last Name");
		document.forms[0].txtLName.select();
	}
	else if(trimTrue(document.forms[0].txtTitle.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter Title");
		document.forms[0].txtTitle.select();
	}
	else if(trimTrue(document.forms[0].txtEmployeeID.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter Employee ID");
		document.forms[0].txtEmployeeID.select();
	}
	else if(trimTrue(document.forms[0].txtDesignation.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter the Designation");
		document.forms[0].txtDesignation.select();
	}
	else if (document.forms[0].selUnit.selectedIndex == "0") {
		alert("Select Organization");
		document.forms[0].selUnit.focus();
		ctr = ctr + 1;
	}
	else if (!document.forms[0].txtWorkEmail.value == "" && !checkmailid(document.forms[0].txtWorkEmail.value,'W')) {
		ctr = ctr + 1;
	}
	else if (!document.forms[0].txtHomeEmail.value == "" && !checkmailid(document.forms[0].txtHomeEmail.value,'H')) {
		ctr = ctr + 1;
	}
	
	if (ctr == 0)
	{
		document.forms[0].submit();
	}
}

function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function checkmailid(mailid,flag) {
	var exclude=/[^@\-\.\w]{2}|[@\.]{2}|(@)[^@]*\1/;
	var check=/@[\w\-]+\./;
	var checkend=/\.[a-zA-Z]{2,3}$/;

	if(((mailid.search(exclude) != -1)||(mailid.search(check))==-1)||(mailid.search(checkend) == -1)) {
		alert("Enter Valid Emailid");
		if (flag == "W")
			document.forms[0].wemail.select();
		else if (flag == "H")
			document.forms[0].hemail.select();
		
		return false;
	}
	else
		return true;
}

