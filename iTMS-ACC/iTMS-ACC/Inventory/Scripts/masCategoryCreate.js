function trimTrue(val) {
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return String(val == null ? "" : val).replace(ltrim, "").replace(rtrim, "");
}

function checkSubmit() {
	var form = document.forms[0];
	if (trimTrue(form.txtCatCode.value) === "") {
		alert("Enter Category Code");
		form.txtCatCode.select();
		return false;
	}
	if (trimTrue(form.txtCatName.value) === "") {
		alert("Enter Category Name");
		form.txtCatName.select();
		return false;
	}
	if (trimTrue(form.txtCatShName.value) === "") {
		alert("Enter Category Short Name");
		form.txtCatShName.select();
		return false;
	}
	form.action = "MasCategoryInsert.asp";
	form.submit();
	return true;
}
