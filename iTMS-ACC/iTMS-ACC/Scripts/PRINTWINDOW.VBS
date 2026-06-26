Function PrintWindow(sPara)
	sWin = showModalDialog(sPara,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")
end Function

Function Paginate(iPageNo)
	document.formname.hPageSelection.value = iPageNo
	document.formname.submit
end Function

Function PaginateAcc(iPageNo)
	document.formname.hPageSelection.value = iPageNo
	GetFormDet
	document.formname.submit
end Function
