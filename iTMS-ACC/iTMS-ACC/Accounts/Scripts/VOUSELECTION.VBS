function popVoucherNo(sCallTy,iBookId,sTrans,sOrgID,sVouType,iBookNo)
	Dim sbTemp,isCTemp
	IF Cstr(sVouType) = "" Then
		sVouType = document.formname.selVouType.Value
	End IF
	
	IF Cstr(sTrans) = "GJR" Then
		IF document.formname.selUnitId.selectedIndex = 0 Then
			Msgbox "Select Unit "
			document.formname.selUnitId.focus()
			Exit Function
		Elseif document.formname.selBook.selectedIndex = 0 Then
			Msgbox "Select Book "
			document.formname.selBook.focus()
			Exit Function
			
		End IF
	Else
		IF document.formname.selUnitId.selectedIndex = 0 Then
			Msgbox "Select Unit "
			document.formname.selUnitId.focus()
			Exit Function
		Elseif document.formname.selBook.selectedIndex = 0 Then
			Msgbox "Select Book "
			document.formname.selBook.focus()
			Exit Function
		Elseif document.formname.selVouType.selectedIndex = 0 Then
			Msgbox "Select Voucher Type "
			document.formname.selVouType.focus()
			Exit Function
		End IF
		
	End IF
	if sVouType="C" then
		sTrans=sTrans&"P"
	elseif sVouType="0" Then
		sTrans = sTrans
	Else
		sTrans=sTrans&"R"
	end if
	
	iscTemp = Split(iBookNo,"-")
	IF UBound(iscTemp) = 1 Then
		iBookNo = iscTemp(0)
	End IF
	
	sTemp = showModalDialog("VoucherSelection.asp?flag="&sCallTy&"&orgId="&sOrgId&"&BookCode="&iBookId&"&BookNo="&iBookNo&"&TransType="&sTrans,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	
	
	
	IF Cstr(sCallTy) = "A" Then
		if Trim(sTemp)<>"0" then
			saTemp=Split(sTemp,"~")			
			document.formname.hTransNo.value=saTemp(0)
			sbTemp=Split(saTemp(1),">>")
			document.formname.txtVouNo.value=sbTemp(0)
			'document.formname.btnAmend.disabled = True
			'document.formname.btnDel.disabled = True
			'document.formname.btnCreate.disabled = True
			document.formname.hAmendTy.value = "A"
			document.formname.selUnitId.disabled = False
			'document.formname.btnDelVou.disabled = True
			MakeDispVou(sCallTy)
		Else
			document.formname.hTransNo.value=""
			document.formname.txtVouNo.value=""
		End IF
		
		'IF Cstr(sTrans) = "GJR" Then
		'	document.formname.btnAmend.disabled = True
		'End IF
	Else
	
		if Trim(sTemp)<>"0" then
			saTemp=Split(sTemp,"~")			
			document.formname.hTransNo.value=saTemp(0)
			document.formname.txtVouNo.value=saTemp(1)
			'document.formname.btnAmend.disabled = false
			'document.formname.btnDel.disabled = false
			document.formname.hAmendTy.value = "N"
			document.formname.selUnitId.disabled = True
			'document.formname.btnDelVou.disabled = False
			MakeDispVou(sCallTy)
			
		else
			document.formname.hTransNo.value=""
			document.formname.txtVouNo.value=""
		end if	
	End IF
	
End function

function popVoucherNoDispOnly(sCallTy,iBookId,sTrans,sOrgID,sVouType,iBookNo)
	Dim sbTemp
	IF Cstr(sTrans) = "GJR" Then
		IF document.formname.selUnitId.selectedIndex = 0 Then
			Msgbox "Select Unit "
			document.formname.selUnitId.focus()
			Exit Function
		Elseif document.formname.selBook.selectedIndex = 0 Then
			Msgbox "Select Book "
			document.formname.selBook.focus()
			Exit Function
			
		End IF
	Else
		IF document.formname.selUnitId.selectedIndex = 0 Then
			Msgbox "Select Unit "
			document.formname.selUnitId.focus()
			Exit Function
		Elseif document.formname.selBook.selectedIndex = 0 Then
			Msgbox "Select Book "
			document.formname.selBook.focus()
			Exit Function
		Elseif document.formname.selVouType.selectedIndex = 0 Then
			Msgbox "Select Voucher Type "
			document.formname.selVouType.focus()
			Exit Function
		End IF
		
	End IF
	if sVouType="C" then
		sTrans=sTrans&"P"
	elseif sVouType="0" Then
		sTrans = sTrans
	Else
		sTrans=sTrans&"R"
	end if
	
	sTemp = showModalDialog("VoucherSelection.asp?flag="&sCallTy&"&orgId="&sOrgId&"&BookCode="&iBookId&"&BookNo="&iBookNo&"&TransType="&sTrans,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	
	IF Cstr(sCallTy) = "A" Then
		if Trim(sTemp)<>"0" then
			saTemp=Split(sTemp,"~")			
			document.formname.hTransNo.value=saTemp(0)
			sbTemp=Split(saTemp(1),">>")
			document.formname.txtVouNo.value=sbTemp(0)
			
		Else
			document.formname.hTransNo.value=""
			document.formname.txtVouNo.value=""
		End IF
		
		
	Else
	
		if Trim(sTemp)<>"0" then
			saTemp=Split(sTemp,"~")			
			document.formname.hTransNo.value=saTemp(0)
			document.formname.txtVouNo.value=saTemp(1)
			
		else
			document.formname.hTransNo.value=""
			document.formname.txtVouNo.value=""
		end if	
	End IF
	
End function

