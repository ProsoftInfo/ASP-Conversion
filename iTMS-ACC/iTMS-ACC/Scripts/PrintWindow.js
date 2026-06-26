function PrintWindow(sPara) {
	window.open(sPara, "PrintWindow", "height=200,width=300,resizable=no,status=no");
}

function Paginate(iPageNo) {
	document.formname.hPageSelection.value = iPageNo;
	document.formname.submit();
}

function PaginateAcc(iPageNo) {
	document.formname.hPageSelection.value = iPageNo;
	if (typeof GetFormDet === "function") {
		GetFormDet();
	}
	document.formname.submit();
}
