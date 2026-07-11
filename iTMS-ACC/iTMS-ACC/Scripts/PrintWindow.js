function PrintWindow(sPara) {
	window.open(sPara, "PrintWindow", "height=200,width=300,resizable=no,status=no");
}

function printForm() {
	return document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
}

function Paginate(iPageNo) {
	var form = printForm();
	if (!form) {
		return;
	}
	form.hPageSelection.value = iPageNo;
	form.submit();
}

function PaginateAcc(iPageNo) {
	var form = printForm();
	if (!form) {
		return;
	}
	form.hPageSelection.value = iPageNo;
	if (typeof GetFormDet === "function") {
		GetFormDet();
	}
	form.submit();
}
