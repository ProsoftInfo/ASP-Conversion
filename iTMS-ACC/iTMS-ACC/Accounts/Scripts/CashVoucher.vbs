FUNCTION  checkFileds()
	Dim iPayNo
	if trim(document.formname.txtNarration.value)="" then
		Msgbox("Enter Narration")
		document.formname.txtNarration.select
		checkFileds=false
		exit Function
	end if
	if ValidateAmount(document.formname.txtAmount.value)=false then
		document.formname.txtAmount.select
		checkFileds=false
		exit Function
	end if
	if CDbl(document.formname.txtAmount.value) > CDbl(dTransLimit) then
		select case sTransFlag
			case "W"
					MsgBox "Amount is greater than the amount limit",,"Warning"
			case "R"
					MsgBox "Amount should be less than "&dTransLimit
					checkFileds=false
					exit Function
		end select
	end if
	
	iSno = 1
	for each HeaderNode in EntryRoot.childNodes
		if HeaderNode.nodeName="PayRec" then
			dAmount=CDbl(document.formname.txtAmount.value)
			dTotalAmtAdjust=0
			iCounter=1
				for each  nodANL in HeaderNode.childNodes
					iCode=nodANL.Attributes.getNamedItem("No").Value
					dTransAmount=nodANL.Attributes.getNamedItem("TransAmount").Value
					dAmtAdjusted=nodANL.Attributes.getNamedItem("AmtAdjusted").Value
					dAmtToAccount=nodANL.Attributes.getNamedItem("AmtToAccount").Value
					sAdjType = nodANL.Attributes.getNamedItem("AdjType").Value
					iPayNo = Trim(nodANL.Attributes.getNamedItem("PayableNo").Value)	

					IF CStr(sAdjType) = "I" Then
						dAmtAdjust=CDbl(dTransAmount)-(CDbl(dAmtAdjusted)+CDbl(dAmtToAccount))
					Else
						dAmtAdjust=CDbl(dTransAmount)-CDbl(dAmtAdjusted)
					End IF
					dTotal=eval("document.formname.txtDocAmount"&iCode&"Z"&iPayNo&"Z"&iSno).value
					dTotal = FormatNumber(dTotal,2,0,0,0)
					dAmtAdjust = FormatNumber(dAmtAdjust,2,0,0,0)
					
					if  CDbl(dTotal)>CDbl(dAmtAdjust) then
						MsgBox """To Adjust Amount"" should be less than ""Document Amount-(Adjusted +To Account)"""
						eval("document.formname.txtDocAmount"&iCode&"Z"&iSno).focus
						checkFileds=false
						exit Function
					else
						dTotalAmtAdjust=CDbl(dTotalAmtAdjust)+CDbl(dTotal)
					end if
					iSno = iSno + 1
				next
				'MsgBox dTotalAmtAdjust & " " & dAmount
				'if  CDbl(dTotalAmtAdjust)>CDbl(dAmount) then
				'	MsgBox "Total of ""To Adjust Amount"" should be less than ""Voucher Amount"""
				'	checkFileds=false
				'	exit Function
				'end if
		end if 'End of Check for PayRec Node
	next
	checkFileds=true
END FUNCTION
'---------------------END OF FUNCTION CHECKFILEDS-------------------------
Function CancelAction(sPage)
	document.formname.action=sPage
	document.formname.submit
end Function
'---------------------End Of Function ActionCancel----------------------------

Function AddNewParty()
	OutValue = showModalDialog("MisParCreate.asp?"&OutValue,"","dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'MsgBox OutValue
	document.formname.txtPayTo.value = OutValue
End Function

Function SelMisParty()
	Dim arrTemp,sRetValue,sParCode,sPartyName,sTemp
	Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
    
    sTempValWindowSize = GetWindowSizeForPopup("4")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)
	
	
	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId,"","dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	IF CStr(OutValue) = "AN" Then
		AddNewParty()
		Exit Function
	End IF
	arrTemp = split(OutValue,":")
	
	while UBound(arrTemp) = 0
		OutValue = showModalDialog("../../Common/"&sProgramName&"?"&OutValue,"","dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
		arrTemp = split(OutValue,":")
	wend	
	
'	OutValue = showModalDialog("../../Common/MisPartySelection.asp","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'	IF CStr(OutValue) = "AN" Then
'		AddNewParty()
'		Exit Function
'	End IF
'	arrTemp = split(OutValue,":")
'
'
'	while UBound(arrTemp) = 0
'		OutValue = showModalDialog("../../Common/MisPartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'		arrTemp = split(OutValue,":")
'	wend

	sRetValue = OutValue
	'MsgBox sRetValue
	if UBound(arrTemp) <= 1 then exit function

	sTemp = Split(sRetValue,":")
	document.formname.txtPayTo.value = sTemp(0)
	'sParTy = sTemp(4)
	'sParSubType = sTemp(3)
	'sParCode = sTemp(1)
	'sPartyName = sTemp(0)
End Function

Function CheckVouStat()
	Dim iTotal,sVouType,sVouName,sAltVouName,sStatus,sCurrDate,sTempCurr,dCurrentBal
	Dim iSelPayRec,iTotPayRec,iRetVal
	
	iSelPayRec = document.formname.hSelPayRecCount.value
	iTotPayRec = document.formname.hPayRecCount.value
	sCurrDate = document.formname.hCurrDate.value
	sVouType = document.formname.hVouCRDR.value
	iTotal = CheckVouAmount()
	IF iEntryNo = 1  and DateDiff("d",document.formname.ctlDate.getDate(),sCurrDate) < 0 Then
		MsgBox "Voucher Date Should be Less than the System Date "
		CheckVouStat = false
		Exit Function
	Else
		CheckVouStat = True
	End IF
	
	IF CDbl(iTotal) < 0 Then
		IF CStr(sVouType) = "C" Then
			MsgBox "Total Voucher Amount is more than the Payment Amount"
		Else
			MsgBox "Total Voucher Amount is more than the Receipt Amount"
		End IF
		CheckVouStat = False
		sStatus = "T"
		Exit Function
	End IF
	
	
	IF CDbl(iTotal) = 0 Then
		MsgBox "Total Voucher amount should be More than Zero "
		CheckVouStat = False
		Exit Function
	Else
		sStatus = "F"
	End IF


	iSelPayRec = CDbl(iSelPayRec)
	iTotPayRec = CDbl(iTotPayRec)
	
	IF iTotPayRec <> 0 and iSelPayRec = 0 Then
		iRetVal = MsgBox("Adjustment is Not made for the Party!!, Continue Without Adjustments? ",4,"Warning")
	End IF
	
	IF CStr(iRetVal) = "7" Then
		CheckVouStat = False
		iEntryNo = Cdbl(iEntryNo - 1)
		Exit Function
	End IF
	
	IF Not CheckAdjVal(iTotal) Then
		iRetVal = MsgBox("Payment Amount is made more than the bill value!!, Continue?  ",4,"Warning")
	End IF
	
	IF CStr(iRetVal) = "7" Then
		CheckVouStat = False
		iEntryNo = Cdbl(iEntryNo - 1)
		Exit Function
	End IF
	
	

	IF CStr(document.formname.hVouCRDR.value) = "C" Then
		sTempCurr = Split(document.all.spCurrBal.innerHtml," ")
		dCurrentBal = Trim(sTempCurr(0))
		'Msgbox dCurrentBal
		'dCurrentBal = Mid(1,dCurrentBal,Len(dCurrentBal)-2)
		'Msgbox dCurrentBal

		dCurrentBal = CDbl(dCurrentBal)
		iTotal = Trim(iTotal)
		'Msgbox iTotal
		iTotal = CDbl(iTotal)
		IF iTotal > dCurrentBal Then
			MsgBox "Voucher Amount is Greater than the Current Balance "
			'document.formname.hTotalAmt.value = CDbl(document.formname.hTotalAmt.value) - CDbl(document.formname.txtAmount.value)
			Exit Function
		End IF
		
		'IF CDbl(iTotal) > 20000 Then
		'	MsgBox "Payment Voucher Amount should not be greater than 20,000 "
		'	CheckVouStat = false
		'	Exit Function
		'End IF
		
	End IF
	
	IF DateDiff("d",document.formname.ctlDate.getDate(),sCurrDate) < 0 Then
		MsgBox "Voucher Date Should be Less than the System Date "
		CheckVouStat = false
		Exit Function
	Else
		CheckVouStat = True
	End IF
	
End Function

Function CheckAdjVal(iTotal)
	Dim sExp,TempNode,iAdjVal,iCtr
	sExp = "//PayRec/Doc"
	iAdjVal = 0
	
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.Length - 1
			iAdjVal = iAdjVal + CDbl(TempNode.item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value)
		Next	
	Else
		CheckAdjVal = True
		Exit Function
	End IF
	
	IF CDbl(iTotal) > CDbl(iAdjVal) Then
		CheckAdjVal = False
		Exit Function
	Else
		CheckAdjVal = True
		Exit Function
	End IF
End Function

Function CheckVouAmount()
	Dim sExp,TempNode,iCount,iRecpTotal,iPayTotal,iRetValue
	
	iRecpTotal = 0
	iPayTotal = 0
	
	'Taking values for the Receipt Amount
	sExp = "//Entry[@CRDR=""C""]"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCount = 0 To TempNode.length - 1
			iRecpTotal = iRecpTotal + Cdbl(TempNode.Item(iCount).Attributes.getNamedItem("Amount").Value)
		Next
	End IF
	
	'Taking values for the Payment Amount
	sExp = "//Entry[@CRDR=""D""]"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCount = 0 To TempNode.length - 1
			iPayTotal = iPayTotal + Cdbl(TempNode.Item(iCount).Attributes.getNamedItem("Amount").Value)
		Next
	End IF
	
	'MsgBox iRecpTotal &" " & iPayTotal &" " & document.formname.selVouType.Value
	
	
	'if the Voucher type is Payment then Payment - Receipt else Receipt - payment
	
	IF CStr(document.formname.hVouCRDR.Value) = "D" Then
	
		iRetValue = CDbl(iRecpTotal) - CDbl(iPayTotal)
	Else
		iRetValue = CDbl(iPayTotal) - CDbl(iRecpTotal)
	End IF
	
	CheckVouAmount = iRetValue
	
End Function

Function SelUnBook()
	Dim iBoolSelCode,iCounter,sTemp2
	iBookSelCode = document.formname.hPreBookSel.Value
	'IF Cstr(iBookSelCode) = "" Then 
	'	iBookSelCode = 1
	'End IF
	document.formname.hTdsNew.value = ""
	'Msgbox iBookSelCode
	'document.formname.selUnitId.selectedIndex = 1
	DisplayBook(document.formname.hOrgId)
	IF document.formname.selBook.length > 1 Then
		IF Cstr(iBookSelCode) = "" Then
			document.formname.selBook.selectedIndex = 1
		Else
			For iCounter = 0 To document.formname.selBook.Length - 1
				sTemp2 = Split(Cstr(document.formname.selBook.Options(iCounter).Value),"-")
				IF Cstr(iBookSelCode) = Cstr(Trim(sTemp2(0))) Then
					document.formname.selBook.selectedIndex = iCounter
					Exit For
				End IF
			Next
		End IF
	Else
		document.formname.selBook.selectedIndex = 0
	End IF
	
	IF document.formname.hCallFrm.Value = "A" Then
		document.formname.hAmendTy.value = "A"
		'document.formname.selUnitId.disabled = False
		'document.formname.btnDelVou.disabled = True
		MakeDispVou("C")
	End IF
	'UpdateXML()
	'IF document.formname.selVouType.value = "D" Then
	'	document.formname.selCRDR(0).checked = True
	'	document.formname.selCRDR(1).disabled = True
	'ElseIF document.formname.selVouType.value = "C" Then
	'	document.formname.selCRDR(1).checked = True
	'	document.formname.selCRDR(0).disabled = True
	'End IF
End Function


Function popVoucSel(sCallTy,sBkCode,sVouTyName,sUnt,sRecTy,sSelBk)
	Dim sEntTy
	IF document.formname.selCRDR(0).checked = True Then
		sEntTy = "D"
	Else
		sEntTy = "C"
	End IF
	'popVoucherNo sCallTy,sBkCode,sVouTyName,sUnt,sEntTy,sSelBk
	'document.formname.selVouType.disabled = True
End Function

Function MakeDispVou(sType)
	Dim iTrNo,objhttp,NodRoot,iBookAcc,iCtr
	document.formname.hAmendTy.value = "Y"
	iTrno = document.formname.hTransNo.value
	 
	'Msgbox iTrNo
	IF Cstr(iTrNo) <> "" Then
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","XMLGetVoucher.asp?TransNo="&iTrNo , false
		objhttp.send
		   ' alert "Chk="&objhttp.responseText
		if objhttp.responseXML.xml <> "" then
			VoucherData.loadXML objhttp.responseXML.xml
			Set NodRoot = VoucherData.documentElement
			iBookAcc = Trim(NodRoot.Attributes.Item(2).Value)
			iBookAcc = iBookAcc&"-"&Trim(NodRoot.Attributes.Item(6).Value)
			
			For iCtr = 0 To document.formname.selBook.length - 1
				'MsgBox Cstr(document.formname.selBook.Options(iCtr).Value)&iBookAcc
				IF Cstr(document.formname.selBook.Options(iCtr).Value) = CStr(iBookAcc) Then
			'		MsgBox iCtr
					document.formname.selBook.selectedIndex = iCtr
					Exit For
				End IF
			Next
			
			IF CStr(sType) = "C" Then
				DisplayVoucher("0") 'With Edit Option  
				document.formname.selCRDR(0).disabled = False
				document.formname.selCRDR(1).disabled = False
				'document.formname.selVouType.disabled = True
				document.formname.btnAdd.disabled = False
				document.formname.selAccHead.disabled = False
				DisplayPayRec()
			Else
				DisplayVoucher("1") 'Without Edit Option  
				document.formname.btnAdd.disabled = True
				document.formname.selAccHead.disabled = True
			End IF
		End IF
	End IF
End Function

Function CheckEntryType(sObj)
	Dim iEntryNo
	IF sObj.selectedIndex = 0 Then
		document.formname.btnNext.disabled = True
		Exit Function
	End IF
	document.formname.btnNext.disabled = False
	
	iEntryNo = document.formname.hEntryNo.value
	IF Cstr(iEntryNo) = "0" or CStr(iEntryNo) = "" Then
		document.formname.hVouCRDR.Value = sObj.Value
		IF Cstr(sObj.Value) = "D" Then
			document.formname.selCRDR(0).checked = True
			document.formname.selCRDR(0).disabled = False
			document.formname.selCRDR(1).disabled = True
		Else
			document.formname.selCRDR(1).checked = True
			document.formname.selCRDR(0).disabled = True
			document.formname.selCRDR(1).disabled = False
		End IF
	Else
		document.formname.selCRDR(0).disabled = False
		document.formname.selCRDR(1).disabled = False
	End IF
			
End Function

Function DisplayPayRec()
	Dim VouRoot,sExp,TempNode
	set VouRoot=VoucherData.documentElement
	sExp = "//Entry"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		document.formname.txtPayTo.value = TempNode.Item(0).Attributes.getNamedItem("Payto").Value
		document.formname.txtPayTo.readOnly = True
	End IF
End Function

Function UpdateXML()
	Dim sExp,Root,TempNode,iTemp,Temparr
	IF document.formname.selBook.selectedIndex <> 0 Then
		Set Root = VoucherData.documentElement
		Temparr = Split(document.formname.selBook.Value,"-")
		'alert("Update="&Root.xml)
		sExp = "//voucher"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			iTemp = Split(document.formname.selBook.value,"-")
			'document.forname.selUnitId.disabled = False
			TempNode.Item(0).Attributes.getNamedItem("UnitNo").Value = document.formname.hOrgId.value
			TempNode.Item(0).Attributes.getNamedItem("UnitName").Value = document.formname.hOrgName.value
			TempNode.Item(0).Attributes.getNamedItem("BookNo").Value = iTemp(0)
			TempNode.Item(0).Attributes.getNamedItem("BookName").Value = document.formname.selBook.options(document.formname.selBook.selectedIndex).Text
			TempNode.Item(0).Attributes.getNamedItem("CRDR").Value = document.formname.hVouCRDR.value
			TempNode.Item(0).Attributes.getNamedItem("VouDate").Value = document.formname.ctlDate.GetDate()
			TempNode.Item(0).Attributes.getNamedItem("BookAcchead").Value = Temparr(1)
		End IF
	End IF 
End Function

Function SetBookAccHead()
	Dim iBook,iAcc,iTemp
	Set Root = UnitBookData.documentElement
	If document.formname.selBook.value = "S" Then
		alert("Select Book")
		document.formname.selBook.focus()
		Exit Function
	End IF
	iTemp = Split(document.formname.selBook.value,"-")
	For Each HeaderNode In Root.childNodes
		if  HeaderNode.Attributes.Item(0).nodeValue=Cstr(iTemp(0)) then
			document.formname.hBookAccHead.value=HeaderNode.Attributes.Item(2).nodeValue
			document.formname.hBookOtherUnit.value=HeaderNode.Attributes.Item(3).nodeValue
			Exit For
		End IF
	Next
	'MsgBox document.formname.hBookAccHead.value
	DisplayBalamt()
End Function

Function DelVouch()
	document.formname.action = "VouDeletion.asp"
	document.formname.submit()
End Function

FUNCTION popAccHead()
	dim iHeadCount

	'iUnitNo=document.formname.hAccUnitId.value
	iUnitNo=document.formname.hOrgId.value
	iHeadCount=cint(document.formname.hHeadCount.value)
	iBkNo=document.formname.hBookcode.value
	document.formname.selAccHead.selectedIndex=0

	for iCounter=1 to iHeadCount
		document.formname.selAccHead.remove(1)
	next

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	'objhttp.Open "GET","XMLGetOrgFreqHeads.asp?BkCode=01&BkNo="&iBkNo&"&orgID=" & iUnitNo , false
	'objhttp.send

	'if objhttp.responseXML.xml <> "" then
	'	OutData.loadXML objhttp.responseXML.xml
	'	Set Root = OutData.documentElement
	'	iCounter=1
'
	'	For Each HeaderNode In Root.childNodes
'
	'		set oText1 = document.createElement("<Option>" )
	'			oText1.Text = HeaderNode.text
	'			oText1.Value = HeaderNode.Attributes.Item(0).nodeValue
'
	'		document.formname.selAccHead.add oText1,iCounter
	'		iCounter=CDbl(iCounter)+1
	'	next
	'		document.formname.hHeadCount.value=CDbl(iCounter)-1
	'		iHeadCount=CDbl(iCounter)+1
	'else
	'	document.formname.hHeadCount.value=0
	'	iHeadCount=2
	'end if

	'for iCounter=iHeadCount+1 to document.formname.selAccHead.length
	'	document.formname.selAccHead.remove(iHeadCount)
	'next

	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
		document.formname.selAccHead.length = 2
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		iCounter=document.formname.selAccHead.length
		For Each HeaderNode In Root.childNodes
			set oText1 = document.createElement("<Option>" )
				oText1.Text = HeaderNode.text
				oText1.Value = HeaderNode.Attributes.Item(0).nodeValue

			document.formname.selAccHead.add oText1,iCounter
			iCounter=CDbl(iCounter)+1
		next
	end if

END FUNCTION

FUNCTION selAccountHead(objAcc)
	
	DIM sVouType,sOrgId,sTemp,iHeadCount,sDesc
	iHeadCount=cint(document.formname.hHeadCount.value)
	if objAcc.selectedIndex >0 then
		if document.formname.hOtherUnitFlag.value=1 then
		'	if document.formname.hAccUnitId.selectedIndex <=0 then
		'		objAcc.selectedIndex=0
		'		document.formname.selAccUnitId.focus
		'	else
				'sOrgId=document.formname.hAccUnitId.value
				sOrgId=document.formname.hOrgId.value
				
				if objAcc.selectedIndex <= iHeadCount then
				
					sTemp=Split(objAcc.value,"?")
					document.formname.hTdsElgi.value = sTemp(4)
				
					IF CStr(sTemp(4)) = "0" Then
						document.formname.txtTdsAmount.disabled = True
						'document.formname.txtTdsper.disabled = True
					Else
						document.formname.txtTdsAmount.disabled = False
						'document.formname.txtTdsper.disabled = False
					End IF
					
					sDesc=objAcc.options(objAcc.selectedIndex).text
					bVouFlag=true
					Set newElem = EntryData.createElement("AccHead")
						newElem.setAttribute "No", trim(sTemp(0))
						newElem.setAttribute "CostCenter", trim(sTemp(1))
						newElem.setAttribute "Analytical", trim(sTemp(2))
						newElem.setAttribute "Name", sDesc
						newElem.setAttribute "Type", "G"
						newElem.setAttribute "TransFalg", trim(sTemp(3))
	   					EntryRoot.appendChild newElem

						sTransFlag=trim(sTemp(3))
						window.spAccHead.innerHTML=sDesc&"&nbsp;"


						document.formname.txtPayto.value = document.formname.hPayto.value


					showCCAnal sOrgId,trim(sTemp(0)),trim(sTemp(1)),trim(sTemp(2))

				elseif objAcc.selectedIndex =iHeadCount+1 then
						
						showGLHead sOrgId
				else
				
					sTemp=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
					showPartyHead  sOrgId,sTemp,document.formname.hVouCRDR.value
				End if 'END OF SELECTED ACCOUNT HEAD TYPE IS GL(1) OR PARTY(>1)
				IF iEntryNo > "0" Then
					document.formname.txtNarration.focus()
					document.formname.txtPayTo.readOnly = True
				Else
					document.formname.txtPayto.focus()
					document.formname.txtPayTo.readOnly = False
				End IF
			'end if	'END OF ACCOUNTING UNIT SELECTED OR NOT
		else

			sOrgId=document.formname.hOrgId.value
			if objAcc.selectedIndex <= iHeadCount then
					sTemp=Split(objAcc.value,"?")
					document.formname.hTdsElgi.value = sTemp(4)
					IF CStr(sTemp(4)) = "0" Then
						document.formname.txtTdsAmount.disabled = True
						'document.formname.txtTdsper.disabled = True
					Else
						document.formname.txtTdsAmount.disabled = False
						'document.formname.txtTdsper.disabled = False
					End IF
					sDesc=objAcc.options(objAcc.selectedIndex).text
					bVouFlag=true
					Set newElem = EntryData.createElement("AccHead")
						newElem.setAttribute "No", trim(sTemp(0))
						newElem.setAttribute "CostCenter", trim(sTemp(1))
						newElem.setAttribute "Analytical", trim(sTemp(2))
						newElem.setAttribute "Name", sDesc
						newElem.setAttribute "Type", "G"
						newElem.setAttribute "TransFalg", trim(sTemp(3))
	   					EntryRoot.appendChild newElem

						sTransFlag=trim(sTemp(3))
						window.spAccHead.innerHTML=sDesc&"&nbsp;"
						document.formname.txtPayto.value = document.formname.hPayTo.value


					showCCAnal sOrgId,trim(sTemp(0)),trim(sTemp(1)),trim(sTemp(2))

				elseif objAcc.selectedIndex =iHeadCount+1 then
						showGLHead sOrgId
				else
					sTemp=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
					showPartyHead  sOrgId,sTemp,document.formname.hVouCRDR.value
				End if 'END OF SELECTED ACCOUNT HEAD TYPE IS GL(1) OR PARTY(>1)
				
			IF iEntryNo > 0 Then
				document.formname.txtNarration.focus()
				document.formname.txtPayTo.readOnly = True
			Else
				document.formname.txtPayTo.focus
				document.formname.txtPayTo.readOnly = False
			End IF
		end if	'END OF BOOK HAS OTHER UNIT TRANSCATION OR NOT CHECK
	End if 'END OF IF ANY ACCOUNT HEAD SELECTED CHECK

END FUNCTION
'---------------------END OF FUNCTION SELACCOUNTHEAD----------------------
FUNCTION showPartyHead(sOrgId,sPartyType,sVouType)

dim sPartyCode,bRecivable,bPayable
dim sDocNo,sInvNo,sInvDate,sAmtRec,sAmtRecd
dim nodAccHead,nodPayRec,nodCC,iSno
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
Dim iPayRecCount,sExp,TempNode,iSelPayRec
Dim sAmtToAdjust,iPayNo,bAdv,sNewNarr
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
set objhttp = CreateObject("Microsoft.XMLHTTP")

	sTempValWindowSize = GetWindowSizeForPopup("12")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)


	set objhttp = CreateObject("Microsoft.XMLHTTP")
    
    OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId&"&Party="&sPartyType,"","dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
    arrTemp = Split(Outvalue,":")
	while UBound(arrTemp)=0
		OutValue = showModalDialog("../../Common/"&sProgramName&"?"&OutValue,"","dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
		arrTemp = Split(Outvalue,":")
	wend

if UBound(arrTemp) <= 1 then 
	document.formname.selAccHead.selectedIndex = 0
	document.formname.selAccHead.focus()
	exit function
End IF

sRetValue = OutValue
sTemp = Split(sRetValue,":")
sParTy = sTemp(4)
sParSubType = sTemp(3)
sParCode = sTemp(1)
sPartyName = sTemp(0)

objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sOrgId&"&ParSubType="&sParSubType&"&ParType=" & sParTy&"&PartyCode="&sParCode , false
objhttp.send

IF objhttp.responseText <> "" Then
	sRetVal2 = objhttp.responseText
	GetPartyHeadXml sParCode,sPartyName,sRetVal2
End IF
Set nodAccHead = AccHeadData.documentElement



'Set nodAccHead = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		bVouFlag=true
		sPartyCode=sPartyType&"?"& HeaderNode.Attributes.Item(0).nodeValue
		HeaderNode.Attributes.Item(0).nodeValue=sPartyCode
		bPayable=HeaderNode.Attributes.Item(1).nodeValue
		bRecivable=HeaderNode.Attributes.Item(2).nodeValue
		bAdv = HeaderNode.Attributes.Item(5).nodeValue

		window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue&"&nbsp;"
		IF document.formname.txtPayTo.value = "" Then
			document.formname.txtPayto.value=HeaderNode.Attributes.Item(3).nodeValue
		End IF
		EntryRoot.appendChild HeaderNode
		sTransFlag="A"
	next
	'if (cint(bRecivable)>=1 and sVouType="D") or (cint(bPayable)>=1 and sVouType="C") then
	if (cint(bRecivable)>=1) or (cint(bPayable)>=1) or (Cint(bAdv)>= 1) then
	sPartyCode = Replace(sPartyCode,"&","and")
	'MsgBox sPartyCode
	'If Selected Party Has Payable or Receiavable
	
		Set nodPayRec = showModalDialog("PayRecSelectionWithAllAdj.asp?orgId="+sOrgId+"&ParCode="+sPartyCode&"&Type="&sVouType,"","")
		'alert(nodPayrec.xml)
		sExp = "//RecCount"
		Set TempNode = nodPayRec.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			iPayRecCount = TempNode.Item(0).Attributes.Item(0).nodeValue
		End IF
		
		sExp = "//Doc"
		Set TempNode = nodPayRec.selectNodes(sExp)
		iSelPayRec =  TempNode.length 
		
		document.formname.hSelPayRecCount.value = iSelPayRec
		document.formname.hPayRecCount.value = iPayRecCount
		
		if nodPayRec.Attributes.Item(0).nodeValue=1 then
			'Set the Additional Display Layer Visible
			For Each HeaderNode In nodPayRec.childNodes
					EntryRoot.appendChild HeaderNode
					if HeaderNode.hasChildNodes then
						'If user has Selected Documnets
						iSno=1
						setPayableDisplay 1
						ClearTable "tblPayable",2,1
						for each  nodCC in HeaderNode.childNodes
							sDocNo=nodCC.Attributes.getNamedItem("No").Value
							sInvNo=nodCC.Attributes.getNamedItem("InvNo").Value
							sInvDate=nodCC.Attributes.getNamedItem("InvDate").Value
							sTransAmount=nodCC.Attributes.getNamedItem("TransAmount").Value
							sAmtAdjusted=nodCC.Attributes.getNamedItem("AmtAdjusted").Value
							sAmtToAccount=nodCC.Attributes.getNamedItem("AmtToAccount").Value
							iPayNo = Trim(nodCC.Attributes.getNamedItem("PayableNo").Value)
							
							sInvNo = Replace(sInvNo,"Advance Receipts","ADV REC")
							sInvNo = Replace(sInvNo,"Advance Payments","ADV PAY")
							sInvNo = Replace(sInvNo,"Purchase Inv","PUR INV")
							sInvNo = Replace(sInvNo,"Sales Inv","SAL INV")
							
							IF InStr(1,Cstr(sInvNo),"Purchase") > 0 Then
								sNewNarr = sNewNarr & Trim(Mid(sInvNo,15)) &" "
							ElseIF InStr(1,Cstr(sInvNo),"SALE") > 0 Then
								sNewNarr = sNewNarr & Trim(Mid(sInvNo,5)) &" "
							ElseIF InStr(1,Cstr(sInvNo),"PUR") > 0 Then
								sNewNarr = sNewNarr & Trim(Mid(sInvNo,4)) &" "
							Else
								sNewNarr = sNewNarr & sInvNo &" "
							End IF
							
							sNewNarr = sNewNarr &","
							
							sTransAmount = CDbl(sTransAmount)
							sAmtAdjusted = CDbl(sAmtAdjusted)
							sAmtToAccount = CDbl(sAmtToAccount)
							
							sAmtToAdjust = Cdbl(sTransAmount - sAmtAdjusted - sAmtToAccount)

							set oRow = document.all.tblPayable.insertRow(iSno+1)
							InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
							InsertCell oRow,1,"",sInvNo,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sInvDate,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",FormatNumber(sTransAmount,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",FormatNumber(sAmtAdjusted,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",FormatNumber(sAmtToAccount,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",FormatNumber(sAmtToAdjust,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,2,"txtDocAmount"&CStr(sDocNo)&"Z"&iPayNo&"Z"&iSno,"0","ExcelInputCell","","",12,10,0,0,"style=""text-align:right"""
							iSno=iSno+1

						next
						'document.formname.txtNarration.Value = sNewNarr
					end if 'End of Check Documnet Node
			next	'End of Processing PayRec Node
		else
			'User Has canceled Documnet Selection
			'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
     		setPayableDisplay 0
		end if	'End of Documnet has Childs Check
    else
		'Selected Head has no Documnets
		'Set the Additional Layer Display Layer Hidden
		setPayableDisplay 0
	end if	'End of Party Has Payable Or Recivables
else
	'User canceled Party Head Selection
	window.spAccHead.innerHTML=""
	document.formname.txtPayto.value=""
	'Set the Additional Layer Display Layer Hidden
	setPayableDisplay 0
end if 'End of Party Head Processing

set nodAccHead=nothing
set nodPayRec=nothing
set nodCC=nothing
END FUNCTION
'---------------------END OF FUNCTION showPartyHead--------------------------
Function showGLHead(sOrgId)
dim iAccCode,bAnal,bCostCenter
dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
dim sCode,sDesc,dRatio,iBookNo,sRetVal,arrTemp,sTemp2,sTdsElgi,sTempVal
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

iBookNo=document.formname.hBookcode.value

'set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?hSelectMode=R&hBal=Y&orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),GLHeadData,"dialogHeight:480px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
'sQuery = OutValue.getAttribute("PassQuery")
'
'if OutValue.getAttribute("Action")="CLOSE" then exit function
'while OutValue.getAttribute("Action")<>"Done" 
'    set	OutValue = showModalDialog("../../Common/GLHeadSelection.asp?"&sQuery,GLHeadData,"dialogHeight:480px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
'	sQuery = OutValue.getAttribute("PassQuery")
'    if OutValue.getAttribute("Action")="CLOSE" then exit function
'wend

sTempValWindowSize = GetWindowSizeForPopup("5")
sArrTempValWindowSize = split(sTempValWindowSize,":")
sProgramName = sArrTempValWindowSize(0)
sPopupHeight = sArrTempValWindowSize(1)
sPopupWidth = sArrTempValWindowSize(2)

set OutValue = showModalDialog("../../Common/"&sProgramName&"?hSelectMode=R&hBal=Y&orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),GLHeadData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
sQuery = OutValue.getAttribute("PassQuery")

if OutValue.getAttribute("Action")="CLOSE" then exit function
while OutValue.getAttribute("Action")<>"Done" 
    set	OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,GLHeadData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	sQuery = OutValue.getAttribute("PassQuery")
    if OutValue.getAttribute("Action")="CLOSE" then exit function
wend


'sRetVal = OutValue
'sTempVal = OutValue

'Msgbox sTempVal
if OutValue.hasChildNodes() then
    For each ndChild in OutValue.childNodes
        sTdsElgi = ndChild.getAttribute("RetField6")
    next
end if

if OutValue.hasChildNodes() then    
    For each ndChild in OutValue.childNodes
        sRetVal =sRetVal & ndChild.getAttribute("RetField0")&":"
        sRetVal =sRetVal & ndChild.getAttribute("RetField1")&":"
        sRetVal =sRetVal & ndChild.getAttribute("RetField2")&":"
        sRetVal =sRetVal & ndChild.getAttribute("RetField3")&":"
        sRetVal =sRetVal & ndChild.getAttribute("RetField4")&":"
        sRetVal =sRetVal & ndChild.getAttribute("RetField5")&":"
        sRetVal =sRetVal & ndChild.getAttribute("RetField6")&":"
        sRetVal =sRetVal & ndChild.getAttribute("RetField7")
    next
end if

document.formname.hTdsElgi.value = sTdsElgi
IF CStr(sTdsElgi) = "1" Then
	document.formname.txtTdsAmount.disabled = False
	'document.formname.txtTdsper.disabled = False
Else
	document.formname.txtTdsAmount.disabled = True
	'document.formname.txtTdsper.disabled = True
End IF
GetGlHeadXml(sRetVal)


Set nodAccHead = AccHeadData.documentElement

'Set nodAccHead = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
'alert(nodAccHead.xml)
if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		bVouFlag=true
		iAccCode=HeaderNode.Attributes.Item(0).nodeValue
		bAnal=HeaderNode.Attributes.Item(1).nodeValue
		bCostCenter=HeaderNode.Attributes.Item(2).nodeValue
		sTransFlag=HeaderNode.Attributes.Item(5).nodeValue
		

		window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue&"&nbsp;"
		
		'document.formname.txtPayto.value=HeaderNode.Attributes.Item(3).nodeValue
		EntryRoot.appendChild HeaderNode
	next
	if Trim(bCostCenter)="" or IsNull(bCostCenter) then bCostCenter = 0 
    if Trim(bAnal)="" or IsNull(bAnal) then bAnal = 0
	if cint(bCostCenter)=1 or cint(bAnal)=1 then
	'If Selected GL Account Head has Cost Center
		Set nodCCAnly = showModalDialog("CCAnalysisSelection.asp?orgId="+sOrgId+"&AccCode="+iAccCode,VoucherData,"")
		if nodCCAnly.Attributes.Item(0).nodeValue=1 then
			'Set the Additional and CCANAL Display Layer Visible
			setADDDisplay 1
			For Each HeaderNode In nodCCAnly.childNodes
				if 	HeaderNode.nodeName="CostCenter" then
					EntryRoot.appendChild HeaderNode
					popCostCenter HeaderNode
				end if 'End of Check for Cost Center Node
				if 	HeaderNode.nodeName="Analytical" then
					EntryRoot.appendChild HeaderNode
					popAnalytical(HeaderNode)
				end if 'End of Check for Analytical Node
			next	'End of Processing CCANAL Node
		else
			'User Has canceled CC,ANAL Selection
			'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
			'blocked and added by ragav on Jan 13,2012
     		' setADDDisplay 0 
     		setADDDisplay 1
		end if	'End of CC,ANAL has Childs Check
    else
		'Selected Head has no CC or ANAL
		'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
		setADDDisplay 0
	end if	'End of GL has Cost Center or not
else
	'User canceled Account Head Selection
	window.spAccHead.innerHTML=""
	document.formname.txtPayto.value=""
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	setADDDisplay 0
end if 'End of GL Head Processing

set nodAccHead=nothing
set nodCCAnly=nothing
set nodCC=nothing
End function
'---------------------End Of Function showGLHead--------------------------
'-----------------------------------
Function AddNew()
    if trim(document.formname.hAction.value)="Edit" then
        AddEntry("U")
    else
        AddEntry("A")
    end if
End Function
'====================================================
Function AddEntry(bFlag)
dim iCode,dRatio,dAmount,sExp,TempNode,iCounter,iTdsPer,sCheckExp,CheckNode
Dim iSno,sNarrChk 
' New Validation for check blank data - included on 02/04/2004
sNarrChk = "N"
'IF document.formname.selAccUnitId.selectedIndex = 0 Then
'	MsgBox "Select Accounting Unit "
'	Exit Function
'End IF

IF Not CheckContraMap() Then 'Checks Each Entry Wheather any Acc Head Mapped for Book and same to Contra.
	Exit Function
End IF
 'alert(VouRoot.xml)
''------------Added on Jan 2nd 2008 by Maheswari------------
IF bFlag = "A" then
	IF VouRoot.hasChildnodes then
		For Each EntNode in VouRoot.childnodes
			IF EntNode.Nodename = "Entry" then
				For each TDSNode in EntNode.childnodes
					'IF TDSNode.NodeName = "TDS" then					
						IF document.formname.SelTDSGrp.selectedIndex <> "0" then
							alert("Only One Entry allowed for TDS")
							exit function
						End IF
					'End IF
				Next
			End IF
		Next
	End IF
End IF 'IF bFlag = "A" then
'------------------------------------------------------------------------------------

IF bFlag = "S" then

	IF Trim(document.formname.txtAmount.value) = "0.00" then
	
		IF CheckVouStat() Then
			IF CheckPayRecAmt() Then
			
				SaveXML
				Exit Function
			Else
				Exit Function
			End IF
		End IF
	end if
end if
' End of Validation
if bVouFlag then
	
	if not checkFileds then exit function
	bSavFlag=true

	if bFlag<>"U" then
		iEntryNo=iEntryNo+1
		EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	end if

	IF CStr(bFlag) <> "S" Then
		document.formname.hTotType.value = "A" 
	End IF
	
	sCheckExp = "//Entry[@No="&iEntryNo&" and @TdsAmount]"
	Set CheckNode = VouRoot.selectNodes(sCheckExp)	

	'This gets checks for all PayTo values in the entry node and updates the same.
	'if bFlag="U" then
		sExp = "//Entry"
		Set Tempnode = VouRoot.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			For iCounter = 0 To TempNode.length - 1
				Tempnode.Item(iCounter).Attributes.getNamedItem("Payto").value = document.formname.txtPayto.value
			Next
		End IF
		document.formname.hPayTo.value = document.formname.txtPayto.value
	'End IF

	VouRoot.Attributes.getNamedItem("VouDate").Value=document.formname.ctlDate.getdate

	if document.formname.selCRDR(0).checked then
		EntryRoot.Attributes.getNamedItem("CRDR").Value=document.formname.selCRDR(0).value
	else
		EntryRoot.Attributes.getNamedItem("CRDR").Value=document.formname.selCRDR(1).value
	end if


	if EntryRoot.Attributes.getNamedItem("CRDR").Value ="C" then
		dTotal=dTotal-CDbl(document.formname.txtAmount.value)
	else
		dTotal=dTotal+CDbl(document.formname.txtAmount.value)
	end if

	EntryRoot.Attributes.getNamedItem("Payto").Value=document.formname.txtPayTo.value
	EntryRoot.Attributes.getNamedItem("Amount").Value=document.formname.txtAmount.value
	
	IF CheckNode.length <> 0 Then
		EntryRoot.Attributes.getNamedItem("TdsAmount").Value=document.formname.txtTdsAmount.value
		'EntryRoot.Attributes.getNamedItem("TdsPercentage").Value = document.formname.txtTdsper.value
		EntryRoot.Attributes.getNamedItem("TDSElgi").Value=document.formname.hTDSElgi.value
	Else
		EntryRoot.setAttribute "TdsAmount",document.formname.txtTdsAmount.value
		'EntryRoot.setAttribute "TdsPercentage",document.formname.txtTdsper.value
		EntryRoot.setAttribute "TDSElgi",document.formname.hTDSElgi.value
	End IF

	if document.formname.hOtherUnitFlag.value=1 then
		'EntryRoot.Attributes.getNamedItem("AccUnit").Value=document.formname.hAccUnitId.value	
		EntryRoot.Attributes.getNamedItem("AccUnit").Value=document.formname.hOrgId.value	
		EntryRoot.Attributes.getNamedItem("AccName").Value=document.formname.hOrgName.value
	else
		EntryRoot.Attributes.getNamedItem("AccUnit").Value=document.formname.hOrgId.value
		EntryRoot.Attributes.getNamedItem("AccName").Value=document.formname.hOrgName.value
	end if
	
	EntryRoot.setAttribute "TDSFlag",""
	
	EntryRoot.setAttribute "GroupId",document.formname.SelTDSGrp.Options(document.formname.SelTDSGrp.SelectedIndex).Value		
	EntryRoot.setAttribute "GroupName",document.formname.SelTDSGrp.Options(document.formname.SelTDSGrp.SelectedIndex).Text
		

	'------------------- Added Newly For TDS by Maheshwari on Mar 1st 2007------------------------
		
			'If Document.formname.SelTDSGrp.selectedIndex <> 0 Then
			'If document.formname.SelTDSGrp.Options(document.formname.SelTDSGrp.SelectedIndex).Value <> 0 then
				set RtTds = TDSData.documentElement
				'alert("1="&RtTDS.xml)
				If RtTDS.haschildnodes then
					For each Entnode in EntryRoot.childnodes
						IF Entnode.nodename  = "TDS" then 
							Set RemNode = Entnode
							EntryRoot.Removechild RemNode
						End IF 
					Next	 
					For each node1 in RtTDS.childnodes
						set Tdsnode = node1
						
						EntryRoot.appendChild Tdsnode
					next 'For each Tdsnode in RtTDS.childnodes
				End If 'If RtTDS.haschildnodes then
				 'alert("2="&EntryRoot.xml)
			'End IF	
		'MsgBox "6"
		'----------------------------------------------------------------------------------------------
		
	IF CStr(bFlag) <> "U" Then
		Set newElem = EntryData.createElement("Narration")
		newElem.text= document.formname.txtNarration.value
		EntryRoot.appendChild newElem
	Else
		for each HeaderNode in EntryRoot.childNodes
			IF CStr(HeaderNode.nodeName) = "Narration" Then
				sNarrChk = "Y"
				HeaderNode.text = document.formname.txtNarration.value
			End IF
		Next
		IF CStr(sNarrChk) <> "Y" Then 'IF Narration Node is Not There create a new.
			Set newElem = EntryData.createElement("Narration")
			newElem.text= document.formname.txtNarration.value
			EntryRoot.appendChild newElem
		End IF
	End IF

	for each HeaderNode in EntryRoot.childNodes
		if 	HeaderNode.nodeName="CostCenter" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				dRatio=eval("document.formname.txtCCRatio"&iCode).value
				dAmount=eval("document.formname.txtCCAmount"&iCode).value
				nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
				nodANL.Attributes.getNamedItem("Amount").Value=dAmount
			next
		end if 'End of Check for Cost Center Node
		if 	HeaderNode.nodeName="Analytical" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value

				dRatio=eval("document.formname.txtANALRatio"&iCode&"Z"&sGroupCode).value
				dAmount=eval("document.formname.txtANALAmount"&iCode&"Z"&sGroupCode).value

				nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
				nodANL.Attributes.getNamedItem("Amount").Value=dAmount
			next
		end if 'End of Check for Analytical Node
		iSno = 1
		if 	HeaderNode.nodeName="PayRec" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				iPayNo = Trim(nodANL.Attributes.getNamedItem("PayableNo").Value)
				dAmount=eval("document.formname.txtDocAmount"&iCode&"Z"&iPayNo&"Z"&iSno).value
				nodANL.Attributes.getNamedItem("AmtToAdjust").Value=dAmount
				iSno = iSno + 1
			next
		end if 'End of Check for Analytical Node

	next
	
	IF Not CheckPayRecAmt() Then
		Exit Function
	End IF

 '====== This is to Insert/append the the entry in same order as on the creation == 
	IF CStr(bFlag) = "U" Then
	 'alert(VouRoot.xml)
	 document.formname.hUpdate.value = "Y"
		Dim iCurrEntNo,insNode,sInsxp
		iCurrEntNo = EntryRoot.Attributes.Item(0).nodeValue
		sInsxp = "//Entry[@No="&iCurrEntNo+1&"]"
		
		Set insNode = VouRoot.selectNodes(sInsxp)
		
		IF insNode.length <> 0 Then
			VouRoot.insertBefore EntryRoot,insNode.Item(0)
		Else
			VouRoot.appendChild EntryRoot	
		End IF
		document.formname.hAction.value = "New"
		document.formname.btnNext.disabled = false
	Else
		VouRoot.appendChild EntryRoot	
	End IF
	'alert("Update="&VouRoot.xml)
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=NewVoucher&Mod=CA", false
	objhttp.send VoucherData.XMLDocument

'====================================================================================

	if bFlag="S" then
		IF CheckVouStat() Then
			SaveXML
			Exit Function
		Else
			document.formname.txtPayTo.readOnly = True
			Exit Function
		End IF
	else
		DisplayVoucher("0")
		clearXML()
		setADDDisplay 0
		setPayableDisplay 0

		'document.formname.txtPayTo.value=""

		document.formname.selCRDR(0).disabled=false
		document.formname.selCRDR(1).disabled=false
		document.formname.selAccHead.selectedIndex = 0
		document.formname.txtNarration.value = ""
		document.formname.txtAmount.value = "0.00"
		document.formname.txtTdsAmount.value = "0.00"
		'document.formname.txtTdsper.value = "0.00"
		document.formname.txtTdsAmount.disabled = True
		'document.formname.txtTdsper.disabled = True
		'document.formname.selVouType.disabled = True

		'document.formname.reset

		sExp = "//Entry"
		Set Tempnode = VouRoot.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			document.formname.hPayTo.value = Tempnode.Item(0).Attributes.getNamedItem("Payto").value
		End IF

		document.formname.btnadd.disabled=false
		document.formname.btnnext.disabled=false
		'document.formname.btnupdate.disabled=true
		'document.formname.btndel.disabled=true
		bEditFlag=true
		bVouFlag=false
	end if
else
	if bFlag="S" then
		IF CheckVouStat() Then
			SaveXML
			Exit Function
		Else
			Exit Function
		End IF
	End if
end if
'alert(VouRoot.xml)
End Function

'---------------------End Of Function AddEntry----------------------------
Function EditEntry(iVouEntryNo,sEditType)
Dim sCheckExp,CheckNode
if sEditType="E" then
    if bEditFlag then
        document.formname.hAction.value = "Edit"
        document.formname.btnNext.disabled = true
        setADDDisplay 0
        setPayableDisplay 0
        bVouFlag=true
        'document.formname.hCallFrm.value = "A"
        document.formname.hTdsNew.value ="Y"
        window.spEntryNo.innerHTML=iVouEntryNo
        sCheckExp = "//Entry[@TdsAmount]"
        Set CheckNode = VouRoot.selectNodes(sCheckExp)
iCCCount=0
iAHCount=0
        'alert(CheckNode.Item(0).xml)
'         alert(VouRoot.xml)	
         sOrgID = VouRoot.getAttribute("UnitNo")
            For Each EntryNode in VouRoot.childNodes
		        if EntryNode.Attributes.Item(0).nodeValue=iVouEntryNo then
			        document.formname.txtAmount.value=EntryNode.Attributes.Item(3).nodeValue
			        if EntryNode.Attributes.Item(1).nodeValue ="C" then
				        document.formname.selCRDR(0).checked=true
			        else
				        document.formname.selCRDR(1).checked=true
			        end if
        			
			        IF Cstr(EntryNode.Attributes.Item(6).nodeValue) <> "" Then
				        'document.formname.txtAmount.value = Cdbl(EntryNode.Attributes.Item(3).nodeValue) + Cdbl(EntryNode.Attributes.Item(6).nodeValue)
				        document.formname.txtAmount.value = Cdbl(EntryNode.Attributes.Item(3).nodeValue)
			        End IF 
			        sAccUnit=EntryNode.Attributes.Item(5).nodeValue 
			        document.formname.txtPayTo.value = EntryNode.Attributes.Item(2).nodeValue
			        document.formname.txtPayTo.readOnly = False
			        'MsgBox EntryNode.Attributes.Item(10).nodeValue 
			        '--------------------Added on  21/03/07 by Maheshwari--------------------------------------
				        'MsgBox Document.formname.SelTDSGrp.length
				        document.formname.SelTDSGrp.disabled = False
				        'MsgBox EntryNode.Attributes.Item(10).nodeValue
				        IF Cstr(EntryNode.Attributes.Item(10).nodeValue) <> "0" Then
					        For iCtr = 0 To Document.formname.SelTDSGrp.length - 1
						        'MsgBox Document.formname.SelTDSGrp.options(iCtr).Value &" == " & EntryNode.Attributes.Item(10).nodeValue  
						        If Document.formname.SelTDSGrp.options(iCtr).Value = EntryNode.Attributes.Item(10).nodeValue  Then
							        Document.formname.SelTDSGrp.value =iCtr
							        Exit For
						        End IF
					        Next
				        End IF
				        'document.formname.SelTDSGrp.options(document.formname.SelTDSGrp.SelectedIndex).Value = EntryNode.Attributes.Item(10).nodeValue
				        'document.formname.SelTDSGrp.options(document.formname.SelTDSGrp.SelectedIndex).Text = EntryNode.GetAttribute("GroupName")
		        '--------------------------------------------------------------------------------------	
			        IF CheckNode.length <> 0 Then
				        document.formname.txtTdsAmount.value = EntryNode.Attributes.Item(6).nodeValue
				        'document.formname.txtTdsper.value = EntryNode.Attributes.Item(8).nodeValue

				        IF CStr(EntryNode.Attributes.Item(7).nodeValue) = "1" Then
					        document.formname.txtTdsAmount.disabled = False
					        'document.formname.txtTdsper.disabled = False
				        Else
					        document.formname.txtTdsAmount.disabled = True
					        'document.formname.txtTdsper.disabled = True
				        End IF
				        document.formname.hTDSElgi.value = EntryNode.Attributes.Item(7).nodeValue
			        Else
				        document.formname.txtTdsAmount.value = "0.00"
				        'document.formname.txtTdsper.value = "0.00"
				        document.formname.hTDSElgi.value = "0"
			        End IF

		        '	if document.formname.hOtherUnitFlag.value=1 then
		        '		for i=1 to document.formname.hAccUnitId.length-1
		        '			if document.formname.hAccUnitId.options(i).value =EntryNode.Attributes.Item(4).nodeValue then
		        '				document.formname.selAccUnitId.selectedIndex=i
		        '			end if
		        '		next
				        popAccHead
		        '	end if
			        sAddtional=""
        			
			        For Each HeaderNode in EntryNode.childNodes
				        if HeaderNode.nodeName="AccHead" then
					        if HeaderNode.Attributes.getNamedItem("Type").value="P" then
						        SelectHead HeaderNode.Attributes.getNamedItem("No").value,"P",document.formname.selAccHead,CInt(document.formname.hHeadCount.value)
					        else
						        SelectHead HeaderNode.Attributes.getNamedItem("No").value,"G",document.formname.selAccHead,CInt(document.formname.hHeadCount.value)
					        end if
					        'document.formname.txtPayTo.value=HeaderNode.Attributes.Item(3).nodeValue
					        window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue
					        bCostCenter = HeaderNode.getAttribute("CostCenter")
					        bAnalytical = HeaderNode.getAttribute("Analytical")
					        nAccCode = HeaderNode.getAttribute("No")
				        end if 'End of Check for Account head Node
				        if 	HeaderNode.nodeName="Narration" then
					        document.formname.txtNarration.value=HeaderNode.text
				        end if 'End of Check for Narration Node

				        if 	HeaderNode.nodeName="CostCenter" then
					        setADDDisplay 1
					        popCostCenter(HeaderNode)
					        iCCCount = 1
				        end if 'End of Check for Cost Center Node

				        if 	HeaderNode.nodeName="Analytical" then
					        setADDDisplay 1
					        popAnalytical(HeaderNode)
					        iAHCount = 1
				        end if 'End of Check for Analytical Node
				        
				        if 	HeaderNode.nodeName="PayRec" then
					        popPayRec(HeaderNode)
				        end if 'End of Check for Analytical Node

			        next 'End of Entry Node Loop
			        'alert(VouRoot.xml)
			    
			    ''added by ragav on Jan 13 ,2012 - if Cash Voucher Creation case CC or AH is not selected means here can select
		        ''begin 
		        if Trim(bCostCenter)="" or IsNull(bCostCenter) then bCostCenter = 0 
                if Trim(bAnalytical)="" or IsNull(bAnalytical) then bAnalytical = 0
			    if iCCCount = 0 or iAHCount = 0 then
			        if cint(bCostCenter)=1 or cint(bAnalytical)=1 then
                        'If Selected GL Account Head has Cost Center
                           Set nodCCAnly = showModalDialog("CCAnalysisSelection.asp?orgId="+sOrgId+"&AccCode="+nAccCode,VoucherData,"")
                            if nodCCAnly.Attributes.Item(0).nodeValue=1 then
	                            'Set the Additional and CCANAL Display Layer Visible
	                            setADDDisplay 1
	                            'alert(nodCCAnly.xml)
	                            For Each ndHeader In nodCCAnly.childNodes
		                            if ndHeader.nodeName="CostCenter" then
			                            EntryNode.appendChild ndHeader
			                            popCostCenter ndHeader
		                            end if 'End of Check for Cost Center Node
		                            if ndHeader.nodeName="Analytical" then
			                            EntryNode.appendChild ndHeader
			                            popAnalytical(ndHeader)
		                            end if 'End of Check for Analytical Node
	                            next	'End of Processing CCANAL Node
                            else
	                            'User Has canceled CC,ANAL Selection
	                            'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	                            setADDDisplay 0
                            end if	'End of CC,ANAL has Childs Check
                '    else
                 '       'Selected Head has no CC or ANAL
                  '      'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
                   '     setADDDisplay 0
                    end if	'End of GL has Cost Center or not
                    
                end if 'if iCCCount = 0 or iAHCount = 0 then
                ''end 
                
                set EntryRoot=VouRoot.removeChild(EntryNode)
			    end if
	        next'End of Voucher Node Loop 
	        
			    
        	'document.formname.btnadd.disabled=true
	        'document.formname.btnnext.disabled=true
	        'document.formname.btnupdate.disabled=false
	        'document.formname.btndel.disabled=false
	        bEditFlag=false
	        bSavFlag=true
	        IF Cstr(iVouEntryNo) = "1" Then
		        document.formname.txtPayTo.readOnly = False
	        Else
		        document.formname.txtPayTo.readOnly = True
	        End IF
    end if 
Else
        setADDDisplay 0
        setPayableDisplay 0
        bVouFlag=true
        document.formname.hTdsNew.value ="Y"
        window.spEntryNo.innerHTML=iVouEntryNo
        sCheckExp = "//Entry[@TdsAmount]"
        Set CheckNode = VouRoot.selectNodes(sCheckExp)
	        For Each EntryNode in VouRoot.childNodes
		        if EntryNode.Attributes.Item(0).nodeValue=iVouEntryNo then
			          set EntryRoot=VouRoot.removeChild(EntryNode)
		        end if
	        next'End of Voucher Node Loop 
	       ' document.formname.btnadd.disabled=true
	       ' document.formname.btnnext.disabled=true
	        'document.formname.btnupdate.disabled=false
	        bEditFlag=false
	        bSavFlag=true
	        IF Cstr(iVouEntryNo) = "1" Then
		        document.formname.txtPayTo.readOnly = False
	        Else
		        document.formname.txtPayTo.readOnly = True
	        End IF
	        DelEntry
End if 'sEditType 
   
End Function
'---------------------End Of Function EditEntry----------------------------
Function DelEntry()
	Dim sVouTy
	sVouTy = document.formname.hVouCRDR.value
	clearXML
	setADDDisplay 0
	setPayableDisplay 0
	DisplayVoucher("0")

	document.formname.txtPayTo.value=""
	window.spEntryNo.innerHTML=iEntryNo

	document.formname.selCRDR(0).disabled=false
	document.formname.selCRDR(1).disabled=false

	'document.formname.reset
	
'	IF CStr(sVouTy) = "D" Then
'		document.formname.selVouType.selectedIndex = 1
'	Else
'		document.formname.selVouType.selectedIndex = 2
'	End IF

	document.formname.btnadd.disabled=false
	document.formname.btnnext.disabled=false
	'document.formname.btnupdate.disabled=true
	'document.formname.btndel.disabled=true
	bVouFlag=false
	bEditFlag=true
	bSavFlag=true
End Function
'---------------------End Of Function DelEntry----------------------------
Function DisplayVoucher(sDispTy)
dim sNarration,sAccount,sAddtional,iSno,sAmount,sCheckExp,CheckNode
dim dTotal,sAccUnit,iTdsAmount,iTdsTotAmount,iTdsPer,idivFixed,idivHeight

set VouRoot=VoucherData.documentElement
iTransNo = VouRoot.getAttribute("CreatedTransNo")
'if Trim(iEntryNo)="" or IsNull(iEntryNo) then iEntryNo = 1
sExp = "//Entry"
set tempNode = vouRoot.selectNodes(sExp)
iEntryCount = tempNode.length
	
idivFixed = "70"
idivHeight = cint(25)
idivHeight = cint(30)* cint(iEntryCount)
idivHeight = cint(idivHeight) + cint(idivFixed)
window.DisVoucher.style.height=idivHeight&"px"
window.DisVoucher.style.visibility="visible"

ClearTable "tblVoucher",1,1
dTotal=0

iEntryNo=0
icounter = 1

sDate=VouRoot.Attributes.Item(5).nodeValue
document.formname.ctlDate.setDate=sDate
 ' alert("Disp="&VouRoot.xml)
 'alert(document.formname.hUpdate.value)
For Each EntryNode in VouRoot.childNodes
	 
	iEntryNo=cint(iEntryNo)+1
	sCheckExp = "//Entry[@No="&iEntryNo&" and @TdsAmount]"
	Set CheckNode = VouRoot.selectNodes(sCheckExp)
'	IF trim(EntryNode.nodename) = "Entry" Then
		EntryNode.Attributes.Item(0).nodeValue=iEntryNo
		'alert "AAA="&(EntryNode.Attributes.Item(3).nodename)
	
		sAmount=FormatNumber(trim(EntryNode.Attributes.Item(3).nodeValue),2,,,0)& "&nbsp;" & EntryNode.Attributes.Item(1).nodeValue
	
		IF CStr(EntryNode.Attributes.Item(1).nodeValue) = "C" Then
			dTotal=dTotal-CDbl(EntryNode.Attributes.Item(3).nodeValue)
		Else
			dTotal=dTotal+CDbl(EntryNode.Attributes.Item(3).nodeValue)
		End IF

		sAccUnit=EntryNode.Attributes.Item(5).nodeValue
	'End IF
	IF CheckNode.length <> 0 Then
		iTdsAmount = EntryNode.Attributes.Item(6).nodeValue
		iTdsPer = EntryNode.Attributes.Item(8).nodeValue
	Else
		iTdsAmount = 0
		iTdsPer = 0
	End IF
	IF iTdsAmount = "" then iTdsAmount = 0
	iTdsTotAmount = iTdsTotAmount + CDbl(iTdsAmount)
	document.formname.hPayTo.value = EntryNode.Attributes.Item(2).nodeValue
	
	' alert(sAmount &"****" &dTotal &"***"&iTdsTotAmount)
	sAddtional=""
	
	For Each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
				if HeaderNode.Attributes.Item(4).nodeValue="P" then
					sAccount=HeaderNode.Attributes.Item(3).nodeValue
				else
					sAccount=HeaderNode.Attributes.Item(0).nodeValue
					sAccount=sAccount& "-" & HeaderNode.Attributes.Item(3).nodeValue
				end if
		end if 'End of Check for Account head Node
		if 	HeaderNode.nodeName="Narration" then
				sNarration=HeaderNode.text
		end if 'End of Check for Narration Node
		if 	HeaderNode.nodeName="CostCenter" then
				for each  nodANL in HeaderNode.childNodes
					sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue&"-"
					sAddtional=sAddtional&nodANL.Attributes.Item(3).nodeValue &"%&nbsp;"
					sAddtional=sAddtional&nodANL.Attributes.Item(4).nodeValue&"<br>"
				next
		end if 'End of Check for Cost Center Node
		if 	HeaderNode.nodeName="Analytical" then
				for each  nodANL in HeaderNode.childNodes
					sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue&"-"
					sAddtional=sAddtional&nodANL.Attributes.Item(3).nodeValue &"%&nbsp;"
					sAddtional=sAddtional&nodANL.Attributes.Item(4).nodeValue&"<br>"
				next
		end if 'End of Check for Analytical Node
		if 	HeaderNode.nodeName="PayRec" then
				for each  nodANL in HeaderNode.childNodes
					sAddtional=sAddtional&nodANL.Attributes.Item(1).nodeValue&":"
					sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue &"-&nbsp;"
					sAddtional=sAddtional&nodANL.Attributes.Item(5).nodeValue&"<br>"
				next
		end if 'End of Check for Analytical Node
	next 'End of Entry Node Loop
	
	iTdsAmount = FormatNumber(iTdsAmount,2,,,0)
	iTdsPer = FormatNumber(iTdsPer,2,,,0)

	
	'-------------------Added on Jan 4th 2008 by Maheswari----------------------------		
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	 
	IF CStr(sDispTy) = "0" Then		
		IF iTransNo <> "" then
			objhttp.Open "GET","XMLGetTDSFlag.asp?TransNo="&iTransNo&"&EntryNo=" & iEntryNo , false
			objhttp.send
			'alert "TDSChk="&objhttp.responsetext
			sRetVal = trim(objhttp.responsetext)
			' alert("Val="&sRetVal)
			IF Cstr(sRetVal) = "" Then
				sRetVal = "N"
			End IF
			
			if sRetVal = trim("Y")  then 
				EXIT FUNCTION
 				'InsertCell oRow,1,"","<a  class=""ExcelDisplayCell""><b>Edit</b></a>","ExcelDisplayCell","Center","top",0,0,0,0,""
			elseif sRetVal = "N" then
				set oRow = document.all.tblVoucher.insertRow(iEntryNo)
				InsertCell oRow,1,"",icounter,"ExcelSerial","Center","top",0,0,0,0,""
				InsertCell oRow,1,"","<img src='../../assets/images/iTMS%20Icons/DeleteIcon.gif' alt='Remove Entry'  onclick=javascript:EditEntry('"&iEntryNo&"','D')>","ExcelDisplayCell","Center","top",0,0,0,0,""
				InsertCell oRow,1,"","<a href=""javascript:EditEntry('"&iEntryNo&"','E')"" class=""ExcelDisplayCell""><b>Edit</b></a>","ExcelDisplayCell","Center","top",0,0,0,0,""
			end if
		Else
				set oRow = document.all.tblVoucher.insertRow(iEntryNo)
				InsertCell oRow,1,"",icounter,"ExcelSerial","Center","top",0,0,0,0,""
				InsertCell oRow,1,"","<img src='../../assets/images/iTMS%20Icons/DeleteIcon.gif' alt='Remove Entry' onclick=javascript:EditEntry('"&iEntryNo&"','D')>","ExcelDisplayCell","Center","top",0,0,0,0,""
				InsertCell oRow,1,"","<a href=""javascript:EditEntry('"&iEntryNo&"','E')"" class=""ExcelDisplayCell""><b>Edit</b></a>","ExcelDisplayCell","Center","top",0,0,0,0,""
			End IF
	Else
		InsertCell oRow,1,"","","ExcelDisplayCell","Center","top",0,0,0,0,""
	End IF
	'InsertCell oRow,1,"",sAccUnit,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAccount,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAddtional,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sNarration,"ExcelDisplayCell","left","top",0,0,0,0,""
	'InsertCell oRow,1,"",FormatNumber(sAmount,2,,,0),"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",iTdsAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",iTdsPer,"ExcelDisplayCell","right","top",0,0,0,0,""

	icounter = icounter + 1

next'End of Voucher Node Loop

	dTotal = FormatNumber(dTotal,2,,,0)
	set oRow = document.all.tblVoucher.insertRow(iEntryNo+1)
	InsertCell oRow,1,"","<b>Total</b>","ExcelDisplayCell","right","top",0,0,6,0,""
	InsertCell oRow,1,"","<input type=""text"" name=""txtTotalAmt"" value="&dTotal&" size=""13"" class=""Formelemread"" style=""text-align:right"" >","ExcelDisplayCell","right","top",0,0,0,0,""
	'InsertCell oRow,1,"iTotalAmt",FormatNumber(dTotal,2,,,0) ,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",FormatNumber(iTdsTotAmount,2,,,0) ,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","" ,"ExcelDisplayCell","right","top",0,0,0,0,"" 
	window.spAccHead.innerHTML=""
	window.spEntryNo.innerHTML=iEntryNo+1

End Function

Function DisplayBalamt()
	Dim objHttp,sTemp,sRetVal,Temparr,iTemp
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	sTemp = document.formname.hOrgId.Value
	iTemp = Split(document.formname.selBook.Value,"-")
	'Msgbox document.formname.selBook.Value
	iBookAccHead = iTemp(1)
	'Msgbox iBookAccHead
	sTemp = sTemp&":"&iBookAcchead
	sTemp = sTemp&":"&document.formname.ctlDate.GetDate

	objhttp.Open "GET","GetDayOpenByDate.asp?sValue="&sTemp , false
	objhttp.send
	'alert objHttp.responseText
	sRetval = objHttp.responseText
	Temparr = Split(sRetVal,"*")
	IF UBound(Temparr) = 1 Then
		'MsgBox Temparr(0)
		'MsgBox Temparr(1)
		IF CDbl(Temparr(0)) >= 0 Then
			document.all.spBookBal.innerHtml = FormatNumber(abs(Temparr(0)),2,,,0) &" Dr "
		Else
			document.all.spBookBal.innerHtml = FormatNumber(abs(Temparr(0)),2,,,0) &" Cr "
		End IF

		IF CDbl(Temparr(1)) >= 0 Then
			document.all.spCurrBal.innerHtml = FormatNumber(abs(Temparr(1)),2,,,0) &" Dr "
		Else
			document.all.spCurrBal.innerHtml = FormatNumber(abs(Temparr(1)),2,,,0) &" Cr "
		End IF
	End IF
End Function

Function CheckContraMap() 
	Dim iBkAccHead,iSelAccHead,sTemp,sExp,TempNode,objhttp,sOrgID,sRetVal
	sTemp = Split(document.formname.selBook.Value,"-")
	iBkAccHead = sTemp(1)
	sExp = "//AccHead[@Type=""G""]"
	sOrgID = document.formname.hOrgId.value
	Set TempNode = EntryRoot.selectNodes(sExp)
	sTemp = ""
	IF TempNode.length <> 0 Then
		iSelAccHead = TempNode.Item(0).Attributes.getNamedItem("No").Value	
		sTemp = iBkAccHead&"-"&iSelAccHead&"-"&sOrgID
		'MsgBox sTemp
		Set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","XMLContraMapChk.asp?sValue="&sTemp , false
		objhttp.send		
		sRetVal = Trim(objhttp.responseText)
		IF Len(sRetVal) = 1 and Cstr(sRetVal) = "F" Then
			MsgBox "Contra Entry Not Defined! Do Relate and then Proceed.  "
			CheckContraMap = False
			Exit Function
		Elseif Len(sRetVal) = 1 and Cstr(sRetVal) = "T" Then		
			CheckContraMap = True
			Exit Function
		Else
			alert objhttp.responseText
			Exit Function
		End IF	
	Else
		CheckContraMap = True
		Exit Function
	End IF
End Function
