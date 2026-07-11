function CheckSubmit() {
	ctr = 0;
	var frm = form();
	var radUserType = frm.elements.radUserType;

	if(trimTrue(frm.txtFName.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter First Name");
		frm.txtFName.select();
	}
	else if (radUserType[0].checked)
	{
		frm.hUserType.value = radUserType[0].value;
		
	}
	else if (radUserType[1].checked)
	{
		frm.hUserType.value = radUserType[1].value;
		
	}
	else if (radUserType[2].checked)
	{
		frm.hUserType.value = radUserType[2].value;
		
	}
	else if(trimTrue(frm.txtLName.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter Last Name");
		frm.txtLName.select();
	}
	else if(trimTrue(frm.txtTitle.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter Title");
		frm.txtTitle.select();
	}
	else if(trimTrue(frm.txtEmployeeID.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter Employee ID");
		frm.txtEmployeeID.select();
	}
	else if(trimTrue(frm.txtDesignation.value) == "")
	{
		ctr = ctr + 1;
		alert("Enter the Designation");
		frm.txtDesignation.select();
	}
	else if (!frm.txtWorkEmail.value == "" && !checkmailid(frm.txtWorkEmail.value,'W')) {
		ctr = ctr + 1;
	}
	else if (!frm.txtHomeEmail.value == "" && !checkmailid(frm.txtHomeEmail.value,'H')) {
		ctr = ctr + 1;
	}

	if (ctr == 0)
	{
		frm.action="AmendUserCreationInsert.asp?sLoginID=" + frm.hLoginID.value + "&sOrgID=" + frm.hOrgID.value  + "&sType=A";
		frm.submit();
	}
}

function CheckDelete(){
	var frm = form();
	frm.action="AmendUserCreationInsert.asp?sLoginID=" + frm.hLoginID.value + "&sOrgID=" + frm.hOrgID.value  + "&sType=D";
	frm.submit();
}

function form() {
	return document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
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
			form().wemail.select();
		else if (flag == "H")
			form().hemail.select();
		
		return false;
	}
	else
		return true;
}

