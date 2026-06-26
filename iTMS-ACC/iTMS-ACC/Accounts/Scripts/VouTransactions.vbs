FUNCTION SelectHead(sAccHead,sType,objHead,iHeadCount)
dim saTemp,iCheckVal
if sType="G" then
	iCheckVal = "F"
	for iCounter=0 to cint(objHead.length)-1
		saTemp = split(objHead.options(icounter).value,"?")
		if saTemp(0)=sAccHead then
			objHead.selectedIndex=iCounter
			iCheckVal = "T"
			exit function
		end if
		
	next
	IF Cstr(iCheckVal) = "F" Then
		for iCounter=0 to cint(objHead.length)-1
			if objHead.options(icounter).value=sAccHead then
				objHead.selectedIndex=iCounter
				iCheckVal = "T"
				exit function
			end if
		Next
	End IF
	objHead.selectedIndex=iHeadCount+1
else
	saTemp=split(sAccHead,"?")
	sAccHead=trim(saTemp(0))&"?"&trim(saTemp(1))
	for iCounter=0 to cint(objHead.length)-1

		if objHead.options(icounter).value=sAccHead then
			objHead.selectedIndex=iCounter
			exit function
		end if		
	next
	objHead.selectedIndex=iHeadCount+1
end if

END FUNCTION

FUNCTION CheckAccHead(nodRoot,sAccHead)
dim sExp
	sExp="//AccHead[@No='"&sAccHead&"']"
	set tempNode=nodRoot.selectNodes(sExp)

	if tempNode.length > 0 then
		CheckAccHead=true
	else
		CheckAccHead=false
	end if
END FUNCTION

FUNCTION popAddAmount()

'dim dAmount,iChildCount,dRatio,dTotal,dRatioTotal,iCounter
Dim iPayNo,iSno


if not checkFileds then 
	document.formname.txtAmount.value=""
	exit function
end if

for each HeaderNode in EntryRoot.childNodes
	
	dim sGroupCode
	if HeaderNode.nodeName="CostCenter" then
		dAmount=CDbl(document.formname.txtAmount.value)
		dTotal=dAmount
		dRatioTotal=0
		iCounter=1
		iChildCount=HeaderNode.childNodes.length
		if Cint(iChildCount)> 0 then
			dRatio=Round(100 / iChildCount,2)
			dAmount= Round((dRatio*dAmount)/100,2)
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value

				if iCounter<iChildCount then
					eval("document.formname.txtCCRatio"&iCode).value=dRatio
					eval("document.formname.txtCCAmount"&iCode).value=dAmount
					nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
					nodANL.Attributes.getNamedItem("Amount").Value=dAmount
					dTotal=CDbl(dTotal)-dAmount
					dRatioTotal=CDbl(dRatioTotal)+dRatio
				else
					eval("document.formname.txtCCRatio"&iCode).value=100-dRatioTotal
					eval("document.formname.txtCCAmount"&iCode).value=dTotal
					nodANL.Attributes.getNamedItem("Ratio").Value=100-dRatioTotal
					nodANL.Attributes.getNamedItem("Amount").Value=dTotal
				end if			
				iCounter=CInt(iCounter)+1			

			next
		end if 'End of Check for Cost Center Child Count
	end if 'End of Check for Cost Center Node
	
	if HeaderNode.nodeName="Analytical" then
	
		dAmount=CDbl(document.formname.txtAmount.value)
		dTotal=dAmount
		dRatioTotal=0
		iCounter=1
		iChildCount=HeaderNode.childNodes.length
		if Cint(iChildCount)> 0 then
			dRatio=Round(100 / iChildCount,2)
			dAmount= Round((dRatio*dAmount)/100,2)
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value
				if iCounter<iChildCount then
					eval("document.formname.txtANALRatio"&iCode&"Z"&sGroupCode).value=dRatio
					eval("document.formname.txtANALAmount"&iCode&"Z"&sGroupCode).value=dAmount
					nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
					nodANL.Attributes.getNamedItem("Amount").Value=dAmount
					dTotal=CDbl(dTotal)-dAmount
					dRatioTotal=CDbl(dRatioTotal)+dRatio
				else
					eval("document.formname.txtANALRatio"&iCode&"Z"&sGroupCode).value=100-dRatioTotal
					eval("document.formname.txtANALAmount"&iCode&"Z"&sGroupCode).value=dTotal
					nodANL.Attributes.getNamedItem("Ratio").Value=100-dRatioTotal
					nodANL.Attributes.getNamedItem("Amount").Value=dTotal
				end if			
				iCounter=CInt(iCounter)+1			

			next
		end if 'End of Check for Analytical Child Count	
	end if 'End of Check for Analytical Node
	
	if HeaderNode.nodeName="PayRec" then
		dAmount=CDbl(document.formname.txtAmount.value)
		dTotal=dAmount
		iCounter=1
		iSno = 1
		iChildCount=HeaderNode.childNodes.length
		if Cint(iChildCount)> 0 then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				dTransAmount=nodANL.Attributes.getNamedItem("TransAmount").Value
				dAmtAdjusted=nodANL.Attributes.getNamedItem("AmtAdjusted").Value
				dAmtToAccount=nodANL.Attributes.getNamedItem("AmtToAccount").Value
				sAdjTy = Trim(nodANL.Attributes.getNamedItem("AdjType").Value)
				iPayNo = Trim(nodANL.Attributes.getNamedItem("PayableNo").Value)
				
				
				IF Cstr(sAdjTy) = "I" Then
					'dAmtAdjust=CDbl(dTransAmount)-(CDbl(dAmtAdjusted)+CDbl(dAmtToAccount))
					dAmtAdjust=CDbl(dTransAmount)- CDbl(dAmtAdjusted) - CDbl(dAmtToAccount)
				Else
					dAmtAdjust=CDbl(dTransAmount)- CDbl(dAmtAdjusted) - CDbl(dAmtToAccount)
				End IF
				
				eval("document.formname.txtDocAmount"&iCode&"Z"&iPayNo&"Z"&iSno).value=FormatNumber(dAmtAdjust,2,,,0)
				
				'if  CDbl(dAmtAdjust)>CDbl(dTotal) then				
				'	eval("document.formname.txtDocAmount"&iCode&"Z"&iPayNo&"Z"&iSno).value=FormatNumber(dTotal,2,,,0)
				'	nodANL.Attributes.getNamedItem("AmtToAdjust").Value=FormatNumber(dTotal,2,,,0)
				'	dTotal=0
				'else
				'	eval("document.formname.txtDocAmount"&iCode&"Z"&iPayNo&"Z"&iSno).value=FormatNumber(dAmtAdjust,2,,,0)
				'	nodANL.Attributes.getNamedItem("AmtToAdjust").Value=FormatNumber(dAmtAdjust,2,,,0)
				'	dTotal=CDbl(dTotal)-dAmtAdjust	
				'end if
				iSno = iSno + 1
			next
		end if 'End of Check for PayRec Child Count		
	end if 'End of Check for PayRec Node		
	
next	

END FUNCTION  
'---------------------END OF FUNCTION popAddAmount-------------------------

Function clearXML()
Set EntryRoot = EntryData.createElement("Entry")
	EntryRoot.setAttribute "No",iEntryNo
	EntryRoot.setAttribute "CRDR","0"
	EntryRoot.setAttribute "Payto","0"
	EntryRoot.setAttribute "Amount","0"
	EntryRoot.setAttribute "AccUnit","0"
	EntryRoot.setAttribute "AccName",""
	EntryRoot.setAttribute "TdsAmount","0"
	EntryRoot.setAttribute "TDSElgi","0"
	EntryRoot.setAttribute "TdsPercentage","0"
	EntryRoot.setAttribute "PayRecAmount","0"

end Function
'---------------------End Of Function clearXML----------------------------

FUNCTION ValidateAmount(dAmount)
if  trim(dAmount)="" then
	Msgbox("Amount Cannot be blank")
	ValidateAmount=false
	exit Function
elseif IsNumeric(dAmount)=false then
	Msgbox("Enter Numeric values for Amount")
	ValidateAmount=false
	exit Function
elseif CDbl(dAmount)<=0 or CDbl(dAmount)>9999999999.99 then
	Msgbox("Amount should be >0 and < 9999999999.99")
	ValidateAmount=false
	exit Function
end if	
ValidateAmount=true
END FUNCTION
'---------------------END OF FUNCTION ValidateAmount--------------------------

FUNCTION showNarration(sBookCode)
dim sOrgId,sBookNo,sNarration,sTemp

sTemp = Split(document.formname.selBook.value,"-")
sOrgId=document.formname.hOrgId.value
'sBookNo=sBookCode&"?"&document.formname.selBook.value
sBookNo=sBookCode&"?"&sTemp(0)

'Msgbox sOrgID

sNarration = showModalDialog("NarrationSelection.asp?orgId="+sOrgId&"&BookCode="&sBookNo,"","")
if sNarration<>"" then document.formname.txtNarration.value=sNarration
End Function
'---------------------END OF FUNCTION showNarration--------------------------

function popCostCenter(HeaderNode)
	if HeaderNode.hasChildNodes then
	'If user has Selected Cost centers
	iSno=1
	setAnalDisplay "C",1
	ClearTable "tblCost",1,1
	for each  nodCC in HeaderNode.childNodes
		sCode=nodCC.Attributes.Item(0).nodeValue
		sDesc=nodCC.Attributes.Item(2).nodeValue
		dRatio=nodCC.Attributes.Item(3).nodeValue
		dAmount=nodCC.Attributes.Item(4).nodeValue			
		set oRow = document.all.tblCost.insertRow(iSno)
		InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
		InsertCell oRow,1,"",sDesc,"ExcelDisplayCell","","",0,0,0,0,""
		InsertCell oRow,2,"txtCCRatio"&CStr(sCode),dRatio,"ExcelInputCell","","",6,5,0,0,""
		InsertCell oRow,2,"txtCCAmount"&CStr(sCode),dAmount,"ExcelInputCell","","",12,10,0,0,""
								
		iSno=iSno+1
	next
	else
		'No Cost Center Selected
		setAnalDisplay "C",0
	end if 'End of Check for Selected Cost centers
End function
'---------------------End Of Function popCostCenter--------------------------
function popAnalytical(HeaderNode)
dim sGroupCode
	if HeaderNode.hasChildNodes then
		iSno=1
		setAnalDisplay "A",1
		ClearTable "tblAnal",1,1
		
		for each  nodANL in HeaderNode.childNodes
			sCode=nodANL.Attributes.Item(0).nodeValue
			sDesc=nodANL.Attributes.Item(2).nodeValue
			dRatio=nodANL.Attributes.Item(3).nodeValue
			dAmount=nodANL.Attributes.Item(4).nodeValue
			sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value

			set oRow = document.all.tblAnal.insertRow(iSno)

			InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
			InsertCell oRow,1,"",sDesc,"ExcelDisplayCell","","",0,0,0,0,""
			InsertCell oRow,2,"txtANALRatio"&CStr(sCode)&"Z"&sGroupCode,dRatio,"ExcelInputCell","","",6,5,0,0,""
			InsertCell oRow,2,"txtANALAmount"&CStr(sCode)&"Z"&sGroupCode,dAmount,"ExcelInputCell","","",12,10,0,0,""
			
			'MsgBox "txtANALRatio"&CStr(sCode)&"Z"&sGroupCode					
			iSno=iSno+1
		next
	else
		'No Analytical Selected
		setAnalDisplay "A",0
	end if 'End of Check for Selected Analytical
End function
'---------------------End Of Function popAnalytical--------------------------

function showCCAnal(sOrgId,iAccCode,bCostCenter,bAnal)
'dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
'dim sCode,sDesc,dRatio,iBookNo
sTransNo = document.formname.hTransNo.Value
sEntNo = document.formname.hEntryNo.value
if cint(bCostCenter)=1 or cint(bAnal)=1 then
'If Selected GL Account Head has Cost Center
	Set nodCCAnly = showModalDialog("CCAnalysisSelection.asp?orgId="+sOrgId+"&AccCode="+iAccCode+"&TransNo="+sTransNo+"&EntNo="+sEntNo,"","")
	'Set nodCCAnly = showModalDialog("CCAnalysisSelection.asp?orgId="+sOrgId+"&AccCode="+iAccCode,"","")
	if nodCCAnly.Attributes.Item(0).nodeValue=1 then
		'Set the Additional and CCANAL Display Layer Visible
		setADDDisplay 1
		For Each HeaderNode In nodCCAnly.childNodes
			
			if 	HeaderNode.nodeName="CostCenter" then
				EntryRoot.appendChild HeaderNode
				if HeaderNode.hasChildNodes then
					'If user has Selected Cost centers
					iSno=1
					setAnalDisplay "C",1
					ClearTable "tblCost",1,1
					for each  nodCC in HeaderNode.childNodes
			
						sCode=nodCC.Attributes.getNamedItem("No").Value
						sDesc=nodCC.Attributes.getNamedItem("ShortName").Value
						dRatio=nodCC.Attributes.getNamedItem("Ratio").Value
							
						set oRow = document.all.tblCost.insertRow(iSno)
						InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sDesc,"ExcelDisplayCell","left","",0,0,0,0,""
						InsertCell oRow,2,"txtCCRatio"&sCode,CStr(dRatio),"ExcelInputCell","","",6,5,0,0,""
						InsertCell oRow,2,"txtCCAmount"&sCode,"0","ExcelInputCell","","",12,10,0,0,""
							
						iSno=iSno+1
					next
				else
					'No Cost Center Selected
					setAnalDisplay "C",0
				end if 'End of Check for Selected Cost centers
			end if 'End of Check for Cost Center Node
			
			if HeaderNode.nodeName="Analytical" then
				
				EntryRoot.appendChild HeaderNode
				if HeaderNode.hasChildNodes then
					iSno=1
					setAnalDisplay "A",1
					ClearTable "tblAnal",1,1
					for each  nodANL in HeaderNode.childNodes
						sCode=nodANL.Attributes.getNamedItem("No").Value
						sDesc=nodANL.Attributes.getNamedItem("ShortName").Value
						dRatio=nodANL.Attributes.getNamedItem("Ratio").Value
						sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value			
										
						set oRow = document.all.tblAnal.insertRow(iSno)
						InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sDesc,"ExcelDisplayCell","","",0,0,0,0,""
						InsertCell oRow,2,"txtANALRatio"&sCode&"Z"&sGroupCode,dRatio,"ExcelInputCell","","",6,5,0,0,""
						InsertCell oRow,2,"txtANALAmount"&sCode&"Z"&sGroupCode,"0","ExcelInputCell","","",12,10,0,0,""
						iSno=iSno+1
					next
				else
					'No Analytical Selected
					setAnalDisplay "A",0
				end if 'End of Check for Selected Analytical
			end if 'End of Check for Analytical Node
		next	'End of Processing CCANAL Node
	else
		'User Has canceled CC,ANAL Selection
		'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
 		setADDDisplay 0
	end if	'End of CC,ANAL has Childs Check 
else
	'Selected Head has no CC or ANAL
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	setADDDisplay 0
end if	'End of GL has Cost Center or not

set nodAccHead=nothing
set nodCCAnly=nothing
set nodCC=nothing

End function
'---------------------End Of Function showCCAnal--------------------------

function popPayRec(HeaderNode)
'dim sDocNo,sInvNo,sInvDate,sAmtRec,sAmtRecd,nodCC,iSno
Dim dTotalAmtToAcc,iPayNo

	if HeaderNode.hasChildNodes then
		'If user has Selected Documnets
		iSno=1
		setPayableDisplay 1
		ClearTable "tblPayable",2,1
		for each  nodCC in HeaderNode.childNodes
			sDocNo=nodCC.Attributes.getNamedItem("No").Value
			sInvNo=nodCC.Attributes.getNamedItem("InvNo").Value
			sInvDate=nodCC.Attributes.getNamedItem("InvDate").Value
			sTransAmount=Trim(nodCC.Attributes.getNamedItem("TransAmount").Value)
			sAmtAdjusted=Trim(nodCC.Attributes.getNamedItem("AmtAdjusted").Value)
			sAmtToAccount=Trim(nodCC.Attributes.getNamedItem("AmtToAccount").Value)
			sAmtToAdjust=nodCC.Attributes.getNamedItem("AmtToAdjust").Value
			iPayNo = Trim(nodCC.Attributes.getNamedItem("PayableNo").Value)
			
			sTransAmount = Cdbl(sTransAmount)
			sAmtAdjusted = Cdbl(sAmtAdjusted)
			sAmtToAccount = Cdbl(sAmtToAccount)
			
			dTotalAmtToAcc = Cdbl(sTransAmount - sAmtAdjusted - sAmtToAccount)
			
			dTotalAmtToAcc = FormatNumber(dTotalAmtToAcc,2,,,0)
			sTransAmount = FormatNumber(sTransAmount,2,,,0)
			sAmtAdjusted = FormatNumber(sAmtAdjusted,2,,,0)
			sAmtToAccount = FormatNumber(sAmtToAccount,2,,,0)
			
			set oRow = document.all.tblPayable.insertRow(iSno+1)
			InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
			InsertCell oRow,1,"",sInvNo,"ExcelDisplayCell","","",0,0,0,0,""
			InsertCell oRow,1,"",sInvDate,"ExcelDisplayCell","","",0,0,0,0,""
			InsertCell oRow,1,"",sTransAmount,"ExcelDisplayCell","Right","",0,0,0,0,""
			InsertCell oRow,1,"",sAmtAdjusted,"ExcelDisplayCell","Right","",0,0,0,0,""
			InsertCell oRow,1,"",sAmtToAccount,"ExcelDisplayCell","Right","",0,0,0,0,""
			InsertCell oRow,1,"",dTotalAmtToAcc,"ExcelDisplayCell","Right","",0,0,0,0,""
			InsertCell oRow,2,"txtDocAmount"&sDocNo&"Z"&iPayNo&"Z"&iSno,sAmtToAdjust,"ExcelInputCell","right","",12,10,0,0,"style=""text-align:right"""
			iSno=iSno+1		
		next
	end if 'End of Check Documnet Node
End function
'---------------------End Of Function popPayRec--------------------------
FUNCTION setPayableDisplay(iFlag)

if iFlag=0 then
	window.Disaddtional.style.height="1px"
	window.Disaddtional.style.visibility="hidden"
	window.DisPayable.style.height="1px"
	window.DisPayable.style.visibility="hidden"
else
	window.Disaddtional.style.height="115px"
	window.Disaddtional.style.visibility="visible"
	window.DisPayable.style.height="110px"
	window.DisPayable.style.visibility="visible"
end if

END FUNCTION
'---------------------END OF FUNCTION SETPAYABLEDISPLAY-------------------
FUNCTION setAnalDisplay(sDisplay,iFlag)
if sDisplay="A" then
	if iFlag=0 then
		window.DisAnal.style.height="1px"
		window.DisAnal.style.width ="1px"
		window.DisAnal.style.visibility="hidden"
	else
		window.DisAnal.style.height="100px"
		window.DisAnal.style.width ="280px"
		window.DisAnal.style.visibility="visible"
	end if
else
	if iFlag=0 then
		window.DisCost.style.height="1px"
		window.DisCost.style.width ="1px"
		window.DisCost.style.visibility="hidden"
	else
		window.DisCost.style.height="100px"
		window.DisCost.style.width ="280px"
		window.DisCost.style.visibility="visible"
	end if
end if	
END FUNCTION
'---------------------END OF FUNCTION SETANALDISPLAY----------------------------
FUNCTION setADDDisplay(iFlag)
	if iFlag=0 then
		window.Disaddtional.style.height="1px"
		window.Disaddtional.style.visibility="hidden"
		window.DisCCANL.style.height="1px"
		window.DisCCANL.style.visibility="hidden"		
	elseif iFlag=2 then
		window.Disaddtional.style.height="115px"
		window.Disaddtional.style.visibility="visible"
	else
		window.Disaddtional.style.height="115px"
		window.Disaddtional.style.visibility="visible"
		window.DisCCANL.style.height="114px"
		window.DisCCANL.style.visibility="visible"
	end if
						
END FUNCTION
'---------------------END OF FUNCTION SETANALDISPLAY----------------------------

'---------------------- Construct GL Head XML ----------------------------------
Function GetGlHeadXml(sValue)
	Dim sTemp,Root,nodRoot
	sTemp = Split(sValue,":")

	Set nodRoot = AccHeadData.documentElement
	
	Set newElem = AccHeadData.createElement("AccHead")
	newElem.setAttribute "No", trim(sTemp(0))
	newElem.setAttribute "CostCenter",trim(sTemp(2))
	newElem.setAttribute "Analytical",trim(sTemp(3))
	newElem.setAttribute "Name",trim(sTemp(5))
	newElem.setAttribute "Type","G"
	newElem.setAttribute "TransFlag",trim(sTemp(4))
	nodRoot.appendChild newElem
End Function

'-----------------------------------------------------------------------------------------

Function GetGlHeadXmlForSalAcc()
	Dim sTemp,Root,nodRoot
	Set nodRoot = AccHeadData.documentElement
	Set newElem = AccHeadData.createElement("AccHead")
	newElem.setAttribute "No", document.formname.hSalAccCode.Value
	newElem.setAttribute "CostCenter","0"
	newElem.setAttribute "Analytical","0"
	newElem.setAttribute "Name",document.formname.hSalAccName.Value
	newElem.setAttribute "Type","G"
	newElem.setAttribute "TransFlag","A"
	nodRoot.appendChild newElem
End Function

'-------------------Get Party Xml -----------------------------

Function GetPartyHeadXml(sCode,sName,sValue2)
	Dim sTemp,Root,nodRoot
	sTemp = Split(sValue2,":")

	Set nodRoot = AccHeadData.documentElement
	
	Set newElem = AccHeadData.createElement("AccHead")
	newElem.setAttribute "No", trim(sCode)
	newElem.setAttribute "Pay",trim(sTemp(0))
	newElem.setAttribute "Rec",trim(sTemp(1))
	newElem.setAttribute "Name",sName
	newElem.setAttribute "Type","P"
	newElem.setAttribute "Adv",trim(sTemp(2))
	nodRoot.appendChild newElem
End Function

'-----------------------------------------------------------------------------------------

Function SetApp(sType)
	IF Cstr(sType) = "Y" Then
		document.formname.selUserId.disabled = False
	Else
		document.formname.selUserId.disabled = True
	End IF
End Function

'==========================================================================================

Function CheckApp()
	IF document.formname.optApprove(0).checked = True and document.formname.selUserId.selectedIndex = 0 Then
		Msgbox "Select Approver "
		document.formname.selUserId.focus()
		CheckApp = False
		Exit Function
	End IF
	
	IF Len(document.formname.txtNarration.Value) > 300 Then
		Msgbox "Narration Should be Less than 300 Characters "
		CheckApp = False
		Exit Function
	Else
		CheckApp = True
		Exit Function
	End IF
	
	IF Cstr(document.formname.selAccUnitId.Value) <> Cstr(document.formname.selUnitId.value) Then
		Msgbox "Created Unit and Accounting Unit is different!!"
		CheckApp = False
		Exit Function
	End IF
		
	
	
End Function

'-------------------------------------------------------------------------------------------------

'==================================================================================================================
Function CheckFinDate()
	Dim sFinFrm,sFinTo,sCurrMonYr,sTemp
	sFinFrm = document.formname.hFinFrm.Value
	sFinTo = document.formname.hFinTo.Value
	sFinFrm = CDbl(sFinFrm)
	sFinTo = CDbl(sFinTo)
	sTemp = document.formname.ctlDate.GetDate()
	sTemp = Split(sTemp,"/")
	sCurrMonYr = sTemp(2)&sTemp(1)
	sCurrMonYr = CDbl(sCurrMonYr)
	
	
	IF sCurrMonYr < sFinFrm Then
		MsgBox "Voucher Date Should Be Between 01/04/"&Left(sFinFrm,4)&" To 31/03/"&Left(sFinTo,4)
		CheckFinDate = False
		Exit Function
	End IF
	
	
	IF sCurrMonYr > sFinTo Then
		MsgBox "Voucher Date Should Be Between 01/04/"&Left(sFinFrm,4)&" To 31/03/"&Left(sFinTo,4)
		CheckFinDate = False
		Exit Function
	End IF
	CheckFinDate = True
End Function
'==========================================================================================================================
function popPayRecAmd(HeaderNode)
'dim sDocNo,sInvNo,sInvDate,sAmtRec,sAmtRecd,nodCC,iSno
Dim dTotalAmtToAcc,iPayNo

	if HeaderNode.hasChildNodes then
		'If user has Selected Documnets
		iSno=1
		setPayableDisplayAmd 1
		ClearTable "tblPayable",2,1
		for each  nodCC in HeaderNode.childNodes
			sDocNo = nodCC.Attributes.getNamedItem("No").Value
			sInvNo = nodCC.Attributes.getNamedItem("InvNo").Value
			sInvDate = nodCC.Attributes.getNamedItem("InvDate").Value
			sTransAmount = Trim(nodCC.Attributes.getNamedItem("TransAmount").Value)
			sAmtAdjusted = Trim(nodCC.Attributes.getNamedItem("AmtAdjusted").Value)
			sAmtToAccount = Trim(nodCC.Attributes.getNamedItem("AmtToAccount").Value)
			sAmtToAdjust = nodCC.Attributes.getNamedItem("AmtToAdjust").Value
			iPayNo = Trim(nodCC.Attributes.getNamedItem("PayableNo").Value)
			
			sInvNo = Replace(sInvNo,"Credit Notes","CR.NO")
			sInvNo = Replace(sInvNo,"Debit Notes","DR.NO")
			sInvNo = Replace(sInvNo,"Advance Receipts","ADV REC")
			sInvNo = Replace(sInvNo,"Advance Payaments","ADV PAY")
			
			IF Instr(1,sInvNo,"SALE") = 0 Then
				sInvNo = sInvNo&" DT:"&sInvDate	
			End IF
			sInvNo = Replace(sInvNo,"SALE INV","SAL")
			
			
			sTransAmount = Cdbl(sTransAmount)
			sAmtAdjusted = Cdbl(sAmtAdjusted)
			sAmtToAccount = Cdbl(sAmtToAccount)
			
			dTotalAmtToAcc = Cdbl(sTransAmount - sAmtAdjusted - sAmtToAccount)
			
			dTotalAmtToAcc = FormatNumber(dTotalAmtToAcc,2,,,0)
			sTransAmount = FormatNumber(sTransAmount,2,,,0)
			sAmtAdjusted = FormatNumber(sAmtAdjusted,2,,,0)
			sAmtToAccount = FormatNumber(sAmtToAccount,2,,,0)
			sAmtToAdjust = FormatNumber(sAmtToAdjust,2,,,0)
			
			set oRow = document.all.tblPayable.insertRow(iSno+1)
			InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
			'InsertCell oRow,3,"chkDelBill"&sDocNo&"Z"&iPayNo&"Z"&iSno,"","ExcelInputCell","Center","",0,0,0,0,""
			InsertCell oRow,1,"",sInvNo,"ExcelDisplayCell","","",0,0,0,0,""
			'InsertCell oRow,1,"",sInvDate,"ExcelDisplayCell","","",0,0,0,0,""
			InsertCell oRow,1,"",sTransAmount,"ExcelDisplayCell","Right","",0,0,0,0,""
			InsertCell oRow,1,"",sAmtAdjusted,"ExcelDisplayCell","Right","",0,0,0,0,""
			InsertCell oRow,1,"",sAmtToAccount,"ExcelDisplayCell","Right","",0,0,0,0,""
			InsertCell oRow,1,"",dTotalAmtToAcc,"ExcelDisplayCell","Right","",0,0,0,0,""
			InsertCell oRow,2,"txtDocAmount"&sDocNo&"Z"&iPayNo&"Z"&iSno,sAmtToAdjust,"ExcelInputCell","right","",12,10,0,0,"style=""text-align:right"""
			iSno=iSno+1		
		next
	end if 'End of Check Documnet Node
End function

FUNCTION setPayableDisplayAmd(iFlag)
	if iFlag=0 then
		window.Disaddtional.style.height="1px"
		window.Disaddtional.style.visibility="hidden"
		window.DisPayable.style.height="1px"
		window.DisPayable.style.visibility="hidden"
	else
		window.Disaddtional.style.height="115px"
		window.Disaddtional.style.visibility="visible"
		window.DisPayable.style.height="90px"
		window.DisPayable.style.visibility="visible"
	end if
End Function

Function ChkEnter()
	IF window.event.keyCode = 13 Then
		window.event.keyCode = 32
	End IF
End Function
