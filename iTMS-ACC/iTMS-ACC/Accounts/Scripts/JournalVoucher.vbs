FUNCTION popAccHead()

	dim iHeadCount

	iUnitNo=document.formname.hAccUnitId.value
	iHeadCount=cint(document.formname.hHeadCount.value)
	iBkNo=document.formname.hBookcode.value
	document.formname.selAccHead.selectedIndex=0

	for iCounter=1 to iHeadCount
		document.formname.selAccHead.remove(1)
	next

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","XMLGetOrgFreqHeads.asp?BkCode=08&BkNo="&iBkNo&"&orgID=" & iUnitNo , false
	objhttp.send


	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		iCounter=1

		For Each HeaderNode In Root.childNodes


			set oText1 = document.createElement("<Option>" )
				oText1.Text = HeaderNode.text
				oText1.Value = HeaderNode.Attributes.Item(0).nodeValue

			document.formname.selAccHead.add oText1,iCounter
			iCounter=CDbl(iCounter)+1
		next
			document.formname.hHeadCount.value=CDbl(iCounter)-1
			iHeadCount=CDbl(iCounter)+1
	else
		document.formname.hHeadCount.value=0
		iHeadCount=2
	end if

	'for iCounter=iHeadCount+1 to document.formname.selAccHead.length
	'	document.formname.selAccHead.remove(iHeadCount)
	'next

	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
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

FUNCTION selAccountHead()

DIM sVouType,sOrgId,sTemp,iHeadCount,sDesc

iHeadCount=cint(document.formname.hHeadCount.value)
if document.formname.hOtherUnitFlag.value=1 then
	
	'	if document.formname.selAccUnitId.selectedIndex <=0 then
	'		MsgBox ("Select Accounting Unit")
	'		document.formname.selAccHead.selectedIndex=0
	'		document.formname.selAccUnitId.focus
	'	else
			sOrgId=document.formname.hAccUnitId.value
			sTemp=Split(document.formname.selAccHead.value,"?")
			
			IF document.formname.selAccHead.selectedIndex = 0 Then
				document.formname.selAccHead.focus()
				Exit Function
			End IF

			'if document.formname.selAccHead.selectedIndex <= iHeadCount then
			 if UBOUND(sTemp) = 4 Then

				'IF Cstr(document.formname.selAccHead.value) <> "G" Then 
					document.formname.hTdsElgi.value = sTemp(4)
					IF CStr(sTemp(4)) = "0" Then
						document.formname.txtTdsAmount.disabled = True
						document.formname.txtTdsper.disabled = True
					Else
						document.formname.txtTdsAmount.disabled = False
						document.formname.txtTdsper.disabled = False
					End IF
				'End IF

				sDesc=document.formname.selAccHead.options(document.formname.selAccHead.selectedIndex).text
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
					'document.formname.txtPayto.value=sDesc

				showCCAnal sOrgId,trim(sTemp(0)),trim(sTemp(1)),trim(sTemp(2))

			'elseif document.formname.selAccHead.selectedIndex =iHeadCount+1 then
			elseif document.formname.selAccHead.value = "G" then
					showGLHead sOrgId
			else
				sTemp=document.formname.selAccHead.value& "?" & document.formname.selAccHead.options(document.formname.selAccHead.selectedIndex).text
				showPartyHead  sOrgId,sTemp,document.formname.hVouCRDR.value
			End if 'END OF SELECTED ACCOUNT HEAD TYPE IS GL(1) OR PARTY(>1)
			document.formname.txtNarration.focus
	'	end if	'END OF ACCOUNTING UNIT SELECTED OR NOT
	else
		sOrgId=document.formname.hOrgId.value
		if document.formname.selAccHead.selectedIndex <= iHeadCount then
				
				sTemp=Split(document.formname.selAccHead.value,"?")
				document.formname.hTdsElgi.value = sTemp(4)
				IF CStr(sTemp(4)) = "0" Then
					document.formname.txtTdsAmount.disabled = True
					document.formname.txtTdsper.disabled = True
				Else
					document.formname.txtTdsAmount.disabled = False
					document.formname.txtTdsper.disabled = False
				End IF

				sDesc=document.formname.selAccHead.options(document.formname.selAccHead.selectedIndex).text
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
					'document.formname.txtPayto.value=sDesc
				'showCCAnal sOrgId,trim(sTemp(0)),trim(sTemp(1)),trim(sTemp(2))

			elseif document.formname.selAccHead.selectedIndex =iHeadCount+1 then
					showGLHead sOrgId
			elseif document.formname.selAccHead.selectedIndex > 0 then
				sTemp=document.formname.selAccHead.value& "?" & document.formname.selAccHead.options(document.formname.selAccHead.selectedIndex).text
				showPartyHead  sOrgId,sTemp,document.formname.hVouCRDR.value
			End if 'END OF SELECTED ACCOUNT HEAD TYPE IS GL(1) OR PARTY(>1)
		document.formname.txtNarration.focus
	end if	'END OF BOOK HAS OTHER UNIT TRANSCATION OR NOT CHECK

END FUNCTION
'---------------------END OF FUNCTION SELACCOUNTHEAD----------------------
FUNCTION showPartyHead(sOrgId,sPartyType,sVouType)
dim sPartyCode,bRecivable,bPayable,sTransAmount
dim sDocNo,sInvNo,sInvDate,sAmtRec,sAmtRecd
dim nodAccHead,nodPayRec,nodCC,iSno
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
Dim iPayRecCount,sExp,TempNode,iSelPayRec,sAmtToAdjust,sEntType
Dim iPayNo
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
    
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
        
 	    If  OutValue<>"" Then
			sRetValue = OutValue
            sTemp = Split(sRetValue,":")
            sParTy = sTemp(4)
            sParSubType = sTemp(3)
            sParCode = sTemp(1)
            sPartyName = sTemp(0)
        end if     
    
    
	objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sOrgId&"&ParSubType="&sParSubType&"&ParType=" & sParTy&"&PartyCode="&sParCode , false
	objhttp.send

	IF objhttp.responseText <> "" Then
		sRetVal2 = objhttp.responseText
		GetPartyHeadXml sParCode,sPartyName,sRetVal2
	End IF
	Set nodAccHead = AccHeadData.documentElement

	if nodAccHead.hasChildNodes then
		'User Has Selected a GL Account Head
		clearXML()
		For Each HeaderNode In nodAccHead.childNodes
			bVouFlag=true
			
			sPartyCode=sPartyType&"?"& HeaderNode.Attributes.Item(0).nodeValue
			HeaderNode.Attributes.Item(0).nodeValue=sPartyCode
			bPayable=HeaderNode.Attributes.Item(1).nodeValue
			bRecivable=HeaderNode.Attributes.Item(2).nodeValue

			window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue&"&nbsp;"
			'document.formname.txtPayto.value=HeaderNode.Attributes.Item(3).nodeValue
			EntryRoot.appendChild HeaderNode
			sTransFlag="A"
		next
		
		IF document.formname.selCRDR(0).checked = True Then
			sVouType = document.formname.selCRDR(1).value
		Else
			sVouType = document.formname.selCRDR(0).value
		End IF
		
		
		if (cint(bRecivable)>=1 and sVouType="D") or (cint(bPayable)>=1 and sVouType="C") then

		'If Selected Party Has Payable or Receiavable
			sPartyCode = Replace(sPartyCode,"&"," and ")
			Set nodPayRec = showModalDialog("PayRecSelection.asp?orgId="+sOrgId+"&ParCode="+sPartyCode&"&Type="&sVouType,"","")
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
								sTransAmount=Trim(nodCC.Attributes.getNamedItem("TransAmount").Value)
								sAmtAdjusted=Trim(nodCC.Attributes.getNamedItem("AmtAdjusted").Value)
								sAmtToAccount=Trim(nodCC.Attributes.getNamedItem("AmtToAccount").Value)
								sAdjTy = Trim(nodCC.Attributes.getNamedItem("AdjType").Value)
								iPayNo = Trim(nodCC.Attributes.getNamedItem("PayableNo").Value)
								
								sTransAmount = CDbl(sTransAmount)
								sAmtAdjusted = CDbl(sAmtAdjusted)
								sAmtToAccount = CDbl(sAmtToAccount)
								
								'MsgBox sTransAmount &" "& sAmtAdjusted &" "& sAmtToAccount
								
								IF CStr(sAdjTy) = "I" Then
									sAmtToAdjust = Cdbl(sTransAmount - sAmtAdjusted - sAmtToAccount)
								Else
									sAmtToAdjust = Cdbl(sTransAmount - sAmtAdjusted)
								End IF

								set oRow = document.all.tblPayable.insertRow(iSno+1)
								InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
								InsertCell oRow,1,"",sInvNo,"ExcelDisplayCell","","",0,0,0,0,""
								InsertCell oRow,1,"",sInvDate,"ExcelDisplayCell","","",0,0,0,0,""
								InsertCell oRow,1,"",FormatNumber(sTransAmount,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
								InsertCell oRow,1,"",FormatNumber(sAmtAdjusted,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
								InsertCell oRow,1,"",FormatNumber(sAmtToAccount,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
								InsertCell oRow,1,"",FormatNumber(sAmtToAdjust,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
								InsertCell oRow,2,"txtDocAmount"&CStr(sDocNo)&"Z"&iPayNo&"Z"&iSno,"0","ExcelInputCell","Right","",12,10,0,0,"style=""text-align:right"""
								iSno=iSno+1
							next
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
		'document.formname.txtPayto.value=""
		'Set the Additional Layer Display Layer Hidden
		setPayableDisplay 0
	end if 'End of Party Head Processing
'Else
	'User canceled Party Head Selection
'	window.spAccHead.innerHTML=""
	'document.formname.txtPayto.value=""
	'Set the Additional Layer Display Layer Hidden
'	setPayableDisplay 0
'end if
set nodAccHead=nothing
set nodPayRec=nothing
set nodCC=nothing
END FUNCTION
'---------------------END OF FUNCTION showPartyHead--------------------------
function showGLHead(sOrgId)
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

'	sRetVal = OutValue
'	sTempVal = OutValue
    
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
	document.formname.txtTdsper.disabled = False
Else
	document.formname.txtTdsAmount.disabled = True
	document.formname.txtTdsper.disabled = True
End IF
GetGlHeadXml(sRetVal)
Set nodAccHead = AccHeadData.documentElement


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
	if cint(bCostCenter)=1 or cint(bAnal)=1 then
	'If Selected GL Account Head has Cost Center
		Set nodCCAnly = showModalDialog("CCAnalysisSelection.asp?orgId="+sOrgId+"&AccCode="+iAccCode,"","")
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
     		setADDDisplay 0
		end if	'End of CC,ANAL has Childs Check
    else
		'Selected Head has no CC or ANAL
		'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
		setADDDisplay 0
	end if	'End of GL has Cost Center or not
else
	'User canceled Account Head Selection
	window.spAccHead.innerHTML=""
	'document.formname.txtPayto.value=""
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	setADDDisplay 0
end if 'End of GL Head Processing

set nodAccHead=nothing
set nodCCAnly=nothing
set nodCC=nothing
End function
'---------------------End Of Function showGLHead--------------------------
Function AddEntry(bFlag)
dim iCode,dRatio,dAmount,sExp,CheckNode,TempNode,iUniIdx,iBkIdx,iPayNo
Dim sTempVouNo,iCtr

'IF document.formname.selAccUnitId.selectedIndex = 0 Then
'	MsgBox "Select Accounting Unit "
'	Exit Function
'End IF

sExp = "//Entry[@TdsAmount]"
Set CheckNode = VouRoot.selectNodes(sExp)
sTempVouNo = document.formname.txtVouNo.value


'iEntryNo =  document.formname.hEditEntNo.value 

 'alert iEntryNo
if bVouFlag then

	if not checkFileds then exit function
	bSavFlag=true

	if bFlag<>"U" then
		 iEntryNo=iEntryNo+1
		EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	end if

	'  Msgbox "VouRoot="& VouRoot.xml
	 ' alert iEntryNo
	sExp = "//Entry[ @No= "&iEntryNo&"  and @TdsAmount]"
	Set CheckNode = EntryRoot.selectNodes(sExp) 

	VouRoot.Attributes.Item(5).nodeValue=document.formname.ctlDate.GetDate
	if document.formname.selCRDR(1).checked then
		EntryRoot.Attributes.Item(1).nodeValue=document.formname.selCRDR(1).value
	else
		EntryRoot.Attributes.Item(1).nodeValue=document.formname.selCRDR(0).value
	end if

	'EntryRoot.Attributes.Item(2).nodeValue=document.formname.txtPayTo.value
	EntryRoot.Attributes.Item(3).nodeValue=document.formname.txtAmount.value

	IF CheckNode.length <> 0 Then
		EntryRoot.Attributes.getNamedItem("TdsAmount").Value=document.formname.txtTdsAmount.value
		EntryRoot.Attributes.getNamedItem("TDSElgi").Value=document.formname.hTdsElgi.value
		EntryRoot.Attributes.getNamedItem("TdsPercentage").Value=document.formname.txtTdsper.value
	Else 
		sExp = "//Entry[@No="&iEntryNo&"]"
		Set TempNode = VouRoot.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).setAttribute "TdsAmount", document.formname.txtTdsAmount.value
			TempNode.Item(0).setAttribute "TDSElgi",document.formname.hTdsElgi.value
			TempNode.Item(0).setAttribute "TdsPercentage",document.formname.txtTdsper.value
		End IF
	End IF

	if document.formname.hOtherUnitFlag.value=1 then
		EntryRoot.Attributes.Item(4).nodeValue=document.formname.hAccUnitId.value
		EntryRoot.Attributes.Item(5).nodeValue=document.formname.hOrgName.value

	else
		EntryRoot.Attributes.Item(4).nodeValue=document.formname.hOrgId.value
		EntryRoot.Attributes.Item(5).nodeValue=document.formname.hOrgName.value
	end if

	Set newElem = EntryData.createElement("Narration")
	newElem.text= document.formname.txtNarration.value
	EntryRoot.appendChild newElem

	for each HeaderNode in EntryRoot.childNodes
		if 	HeaderNode.nodeName="CostCenter" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.Item(0).nodeValue
				dRatio=eval("document.formname.txtCCRatio"&iCode).value
				dAmount=eval("document.formname.txtCCAmount"&iCode).value
				nodANL.Attributes.Item(3).nodeValue=dRatio
				nodANL.Attributes.Item(4).nodeValue=dAmount
			next
		end if 'End of Check for Cost Center Node
		if 	HeaderNode.nodeName="Analytical" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.Item(0).nodeValue
				sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value

				dRatio=eval("document.formname.txtANALRatio"&iCode&"Z"&sGroupCode).value
				dAmount=eval("document.formname.txtANALAmount"&iCode&"Z"&sGroupCode).value


				nodANL.Attributes.Item(3).nodeValue=dRatio
				nodANL.Attributes.Item(4).nodeValue=dAmount
			next
		end if 'End of Check for Analytical Node
		if 	HeaderNode.nodeName="PayRec" then
			iCtr = 1
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.Item(0).nodeValue
				iPayNo = Trim(nodANL.Attributes.getNamedItem("PayableNo").Value)
				
				dAmount=eval("document.formname.txtDocAmount"&iCode&"Z"&iPayNo&"Z"&iCtr).value
				nodANL.Attributes.Item(5).nodeValue=dAmount
				iCtr = Cdbl(iCtr) + 1
			next
		end if 'End of Check for Analytical Node

	next
'====== This is to Insert/append the the entry in same order as on the creation ==
'alert "b4="&EntryRoot.xml
	IF CStr(bFlag) = "U" Then
		Dim iCurrEntNo,insNode,sInsxp
		iCurrEntNo = EntryRoot.Attributes.Item(0).nodeValue
		'sInsxp = "//Entry[@No="&iCurrEntNo+1&"]"		 
		sInsxp = "//Entry[@No="&iCurrEntNo&"]"		 
		Set insNode = VouRoot.selectNodes(sInsxp)
		IF insNode.length <> 0 Then
			VouRoot.insertBefore EntryRoot,insNode.Item(0)
		Else
			VouRoot.appendChild EntryRoot
		End IF
		document.formname.hAction.value ="New"
		document.formname.btnNext.disabled = false
	Else
		VouRoot.appendChild EntryRoot
	End IF
	'alert "After="&EntryRoot.xml
	'exit function
'====================================================================================


	if bFlag="S" then
		IF CheckVouStat() Then
			IF CheckAmount() Then
				SaveXML
				Exit Function
			End IF
		Else
			Exit Function
		End IF
		Exit Function
	else
		DisplayVoucher("0")
		clearXML()
		document.formname.txtTdsAmount.value = "0.00"
		document.formname.txtTdsper.value = "0.00"
		setADDDisplay 0
		setPayableDisplay 0

		'document.formname.txtPayTo.value=""
		document.formname.selCRDR(0).disabled=false
		document.formname.selCRDR(1).disabled=false
		'iUniIdx = document.formname.hOrgId.value
		iBkIdx = document.formname.selBook.selectedIndex

		document.formname.reset
		document.formname.txtVouNo.value = sTempVouNo

		'document.formname.selUnitId.selectedIndex = iUniIdx
		document.formname.selBook.selectedIndex = iBkIdx


		'document.formname.btnadd.disabled=false
		document.formname.btnnext.disabled=false
		'document.formname.btnupdate.disabled=true
		'document.formname.btndel.disabled=true
		bEditFlag=true
		bVouFlag=false
	end if
else
	if bFlag="S" then
		if not CheckAmount then exit function
		IF CheckVouStat() Then
			SaveXML
			Exit Function
		Else
			Exit Function
		End IF
		Exit Function
	End if
end if

End Function

'---------------------End Of Function AddEntry----------------------------
Function EditEntry(iVouEntryNo,sEditType)
Dim sExp,CheckNode
sExp = "//Entry[@TdsAmount]"
Set CheckNode = VouRoot.selectNodes(sExp)

if sEditType="E" then
    document.formname.hAction.value ="Edit"
    document.formname.btnNext.disabled = true
'alert VouRoot.xml
    if bEditFlag then
        setADDDisplay 0
        setPayableDisplay 0
        bVouFlag=true
        window.spEntryNo.innerHTML=iVouEntryNo

	    For Each EntryNode in VouRoot.childNodes
    		
		    if EntryNode.Attributes.Item(0).nodeValue=iVouEntryNo then
			    'alert iVouEntryNo
			    'document.formname.hEditEntNo.value  = iVouEntryNo
			    document.formname.txtAmount.value=EntryNode.Attributes.Item(3).nodeValue

			    IF CheckNode.length <> 0 Then
				    document.formname.txtTdsAmount.value = EntryNode.Attributes.Item(6).nodeValue
				    document.formname.txtTdsper.value = EntryNode.Attributes.Item(8).nodeValue

				    IF CStr(EntryNode.Attributes.Item(7).nodeValue) = "1" Then
					    document.formname.txtTdsAmount.disabled = False
					    document.formname.txtTdsper.disabled = False
				    Else
					    document.formname.txtTdsAmount.disabled = True
					    document.formname.txtTdsper.disabled = True
				    End IF
				    document.formname.hTDSElgi.value = EntryNode.Attributes.Item(7).nodeValue
			    Else
				    document.formname.txtTdsAmount.value = "0.00"
				    document.formname.txtTdsper.value = "0.00"
				    document.formname.hTDSElgi.value = "0"
			    End IF

			    if EntryNode.Attributes.Item(1).nodeValue ="C" then
					    document.formname.selCRDR(1).checked=true
			    else
					    document.formname.selCRDR(0).checked=true
			    end if

			    sAccUnit=EntryNode.Attributes.Item(5).nodeValue

			    if document.formname.hOtherUnitFlag.value=1 then
			    '	for i=1 to document.formname.selAccUnitId.length-1
			    '		if document.formname.selAccUnitId.options(i).value =EntryNode.Attributes.Item(4).nodeValue then
			    '			document.formname.selAccUnitId.selectedIndex=i
			    '		end if
			    '	next
				    popAccHead
			    end if
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
				    end if 'End of Check for Account head Node
				    if 	HeaderNode.nodeName="Narration" then
					    document.formname.txtNarration.value=HeaderNode.text
				    end if 'End of Check for Narration Node

				    if 	HeaderNode.nodeName="CostCenter" then
					    setADDDisplay 1
					    popCostCenter(HeaderNode)
				    end if 'End of Check for Cost Center Node

				    if 	HeaderNode.nodeName="Analytical" then
					    setADDDisplay 1
					    popAnalytical(HeaderNode)
				    end if 'End of Check for Analytical Node

				    if 	HeaderNode.nodeName="PayRec" then
					    popPayRec(HeaderNode)
				    end if 'End of Check for Analytical Node

			    next 'End of Entry Node Loop
			    set EntryRoot=VouRoot.removeChild(EntryNode)
		    end if
	    next'End of Voucher Node Loop
	    'alert "Test="& VouRoot.xml
	   ' document.formname.btnadd.disabled=true
	    'document.formname.btnnext.disabled=true
	    'document.formname.btnupdate.disabled=false
	    bEditFlag=false
	    bSavFlag=true
    end if
else ' if sEditType="E" then
        setADDDisplay 0
        setPayableDisplay 0
        bVouFlag=true
        window.spEntryNo.innerHTML=iVouEntryNo

	    For Each EntryNode in VouRoot.childNodes
		    if EntryNode.Attributes.Item(0).nodeValue=iVouEntryNo then
			    set EntryRoot=VouRoot.removeChild(EntryNode)
		    end if
	    next'End of Voucher Node Loop
	   
	    bEditFlag=false
	    bSavFlag=true
	    DelEntry()
end if 'if sEditType="E" then
End Function
'---------------------End Of Function EditEntry----------------------------
Function DelEntry()
	clearXML
	
	setADDDisplay 0
	setPayableDisplay 0
	DisplayVoucher("0")
	
	'document.formname.txtPayTo.value=""

	document.formname.selCRDR(0).disabled=false
	document.formname.selCRDR(1).disabled=false

	document.formname.reset

	DisplayBook()
	
	'document.formname.btnadd.disabled=false
	document.formname.btnnext.disabled=false
	'document.formname.btnupdate.disabled=true
	'document.formname.btndel.disabled=true



	bVouFlag=false
	bEditFlag=true
	bSavFlag=true
End Function
'---------------------End Of Function DelEntry----------------------------
FUNCTION CheckAmount()
dim dCrTotal,dDrTotal
dCrTotal=0
dDrTotal=0
For Each EntryNode in VouRoot.childNodes
	sAmount=EntryNode.Attributes.Item(3).nodeValue
	if EntryNode.Attributes.Item(1).nodeValue ="C" then
		dCrTotal=CDbl(dCrTotal)+CDbl(sAmount)
	else
		dDrTotal=CDbl(dDrTotal)+CDbl(sAmount)
	end if
next'End of Voucher Node Loop

dCrTotal = Trim(dCrTotal)
dDrTotal = Trim(dDrTotal)
dCrTotal = Cdbl(dCrTotal)
dDrTotal = Cdbl(dDrTotal)

if dCrTotal<>dDrTotal then
	MsgBox "Cr Total should be equal to Dr Total"
	CheckAmount=false
	exit function
end if
CheckAmount=true
END FUNCTION
'---------------------END OF FUNCTION CHECKAMOUNT----------------------
Function DisplayVoucher(sDispTy)
dim sNarration,sAccount,sAddtional,iSno,sAmount,dTdsAmt,dTdsPer,sExp,CheckNode
dim dTotal,sAccUnit,dCrTotal,dDrTotal,idivFixed,idivHeight

set VouRoot=VoucherData.documentElement
'alert(iEntryNo)
idivFixed = "70"
idivHeight = cint(40)* cint(iEntryNo)
idivHeight = cint(idivHeight) + cint(idivFixed)
window.DisVoucher.style.height=idivHeight&"px"
window.DisVoucher.style.visibility="visible"

ClearTable "tblVoucher",1,1

dTotal=0
dCrTotal=0
dDrTotal=0
iEntryNo=0

sDate=VouRoot.Attributes.Item(5).nodeValue
document.formname.ctlDate.setDate=sDate
For Each EntryNode in VouRoot.childNodes
	iEntryNo=cint(iEntryNo)+1

	sExp = "//Entry[@No="&iEntryNo&" and @TdsAmount]"
	Set CheckNode = VouRoot.selectNodes(sExp)

	EntryNode.Attributes.Item(0).nodeValue=iEntryNo
	sAmount=FormatNumber(EntryNode.Attributes.Item(3).nodeValue,2,,,0)
	'MsgBox EntryRoot.xml
	IF CheckNode.length <> 0 Then
		dTdsAmt = cdbl(EntryNode.Attributes.Item(6).nodeValue)
		dTdsPer = cdbl(EntryNode.Attributes.Item(8).nodeValue)
	Else
		dTdsAmt = 0
		dTdsPer = 0
	End IF

	dTdsAmt = FormatNumber(dTdsAmt,2,,,0)
	dTdsPer = FormatNumber(dTdsPer,2,,,0)

	dTotal=dTotal+CDbl(EntryNode.Attributes.Item(3).nodeValue)
	sAccUnit=EntryNode.Attributes.Item(5).nodeValue
	if EntryNode.Attributes.Item(1).nodeValue ="C" then
		dTotal=dTotal-CDbl(sAmount)
		dCrTotal=CDbl(dCrTotal)+CDbl(sAmount)
	else
		dTotal=dTotal+CDbl(sAmount)
		dDrTotal=CDbl(dDrTotal)+CDbl(sAmount)
	end if

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

	set oRow = document.all.tblVoucher.insertRow(iEntryNo)
	InsertCell oRow,1,"",iEntryNo,"ExcelSerial","Center","top",0,0,0,0,""
	IF CStr(sDispTy) = "0" Then
	    InsertCell oRow,1,"","<img src='../../assets/images/iTMS%20Icons/DeleteIcon.gif' alt='Remove Entry' onclick=javascript:EditEntry('"&iEntryNo&"','D')>","ExcelDisplayCell","Center","top",0,0,0,0,""
		InsertCell oRow,1,"","<a href=""javascript:EditEntry('"&iEntryNo&"','E')"" class=""ExcelDisplayCell""><b>Edit</b></a>","ExcelDisplayCell","Center","top",0,0,0,0,""
	Else
	    InsertCell oRow,1,"","","ExcelDisplayCell","Center","top",0,0,0,0,""
		InsertCell oRow,1,"","","ExcelDisplayCell","Center","top",0,0,0,0,""
	End IF
	'InsertCell oRow,1,"",sAccUnit,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAccount,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sNarration,"ExcelDisplayCell","left","top",0,0,0,0,""
	if EntryNode.Attributes.Item(1).nodeValue ="C" then
		InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
		InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	else
		InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
		InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
	end if

	InsertCell oRow,1,"",sAddtional,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",dTdsAmt,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",dTdsPer,"ExcelDisplayCell","right","top",0,0,0,0,""


next'End of Voucher Node Loop
	dCrTotal=FormatNumber(dCrTotal,2,,,0)
	dDrTotal=FormatNumber(dDrTotal,2,,,0)

	set oRow = document.all.tblVoucher.insertRow(iEntryNo+1)
	InsertCell oRow,1,"","<b>Total</b>","ExcelDisplayCell","right","top",0,0,5,0,""
	InsertCell oRow,1,"",CStr(dCrTotal),"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",CStr(dDrTotal),"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""

	window.spAccHead.innerHTML=""
	window.spEntryNo.innerHTML=iEntryNo+1
	document.formname.hEntryNo.value =iEntryNo+1
End Function

Function CheckVouStat
	Dim sCurrDate
	sCurrDate = document.formname.hCurrDate.value
	IF DateDiff("d",document.formname.ctlDate.getDate(),sCurrDate) < 0 Then
		MsgBox "Voucher Date Should be Less than the System Date "
		CheckVouStat = false
		Exit Function
	Else
		CheckVouStat = True
	End IF	
End Function

'---------------------End Of Function SaveXML-----------------------------
FUNCTION  checkFileds()
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
	checkFileds=true
END FUNCTION
'---------------------END OF FUNCTION CHECKFILEDS-------------------------
Function CancelAction(sPage)
	document.formname.action=sPage
	document.formname.submit
end Function
'===========================================
'-----------------------------------
Function AddNew()
    if trim(document.formname.hAction.value)="Edit" then
        AddEntry("U")
    else
        AddEntry("A")
    end if
End Function
'---------------------End Of Function ActionCancel----------------------------
Function AddNewParty()
	OutValue = showModalDialog("MisParCreate.asp?"&OutValue,"","dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'MsgBox OutValue
	'document.formname.txtPayTo.value = OutValue
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
'	while UBound(arrTemp) = 0
'		OutValue = showModalDialog("../../Common/MisPartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'		arrTemp = split(OutValue,":")
'	wend

	sRetValue = OutValue
	'MsgBox sRetValue
	if UBound(arrTemp) <= 1 then exit function

	sTemp = Split(sRetValue,":")
	'document.formname.txtPayTo.value = sTemp(0)
	'sParTy = sTemp(4)
	'sParSubType = sTemp(3)
	'sParCode = sTemp(1)
	'sPartyName = sTemp(0)
End Function

Function DisplayBook()
dim iUnitNo,arrTemp
dim Root
document.formname.selBook.options.length = 1
'document.formname.selAccUnitId.selectedIndex = objUnit.selectedIndex
'document.formname.hOrgId.value = document.formname.selAccUnitId.Value
'document.formname.hOrgName.value = document.formname.selAccUnitId.Options(document.formname.selAccUnitId.selectedIndex).text

'if objUnit.selectedIndex <> "0" then
	iUnitNo= document.formname.hOrgId.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=08&orgID=" & iUnitNo , false
	objhttp.send
	'alert objhttp.responseXML.xml
	if objhttp.responseXML.xml <> "" then
		UnitBookData.loadXML objhttp.responseXML.xml
		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			document.formname.selBook.length = document.formname.selBook.length+1
			document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
			document.formname.selBook.options(document.formname.selBook.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
		next
	end if
'end if
IF document.formname.selBook.length > 1 Then
	document.formname.selBook.selectedIndex = 1
End IF
End Function

Function SetUnBook()
	'document.formname.selUnitId.selectedIndex = 1
	DisplayBook()
	IF document.formname.selBook.length > 1 Then
		document.formname.selBook.selectedIndex = 1
	Else
		document.formname.selBook.selectedIndex = 0
	End IF
	 
	IF document.formname.hCallFrm.Value = "A" Then
		document.formname.hAmendTy.value = "A"
		'document.formname.selUnitId.disabled = False
		'document.formname.btnDelVou.disabled = True
 
		MakeDispVou("C")
	End IF
	
	'Msgbox "1"
	'document.formname.ctlDate.setDate = document.formname.sMaxDate.Value
	'IF document.formname.selUnitId.length > 1 Then
	'	document.formname.selUnitId.selectedIndex = 1
	'	DisplayBook(document.formname.selUnitId)
	'	IF document.formname.selBook.length > 1 Then
	'		document.formname.selBook.selectedIndex = 1
	'	End IF
	'Else
	'	Msgbox "Error in Populating Unit "
	'End IF
	
End Function

Function MakeDispVou(sType)
	Dim iTrNo,objhttp,sVouNo,sExp,TempNode,TempRoot,iCtr
	document.formname.hAmendChk.value = "Y"
	iTrno = document.formname.hTransNo.value
	 
	IF Cstr(iTrNo) <> "" Then
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","XMLGetVoucher.asp?TransNo="&iTrNo , false
		objhttp.send
		'alert objhttp.responsetext
		' alert "XML="&objhttp.responseXML.xml
		if objhttp.responseXML.xml <> "" then
			VoucherData.loadXML objhttp.responseXML.xml
			Set TempRoot = VoucherData.documentElement
			sExp = "//voucher"
			Set TempNode = TempRoot.selectNodes(sExp)
			IF TempNode.Length <> 0 Then
				sVouNo = TempNode.Item(0).Attributes.getNamedItem("VoucherNo").Value
				document.formname.hPreBookNo.value = TempNode.Item(0).Attributes.getNamedItem("BookNo").Value
			End IF
			document.formname.txtVouNo.Value = sVouNo
			IF CStr(sType) = "C" Then
				DisplayVoucher("0") 'With Edit Option
				'document.formname.btnAdd.disabled = False
				'document.formname.btnVouDel.disabled = False
				document.formname.selAccHead.disabled = False
				'DisplayPayRec()
			Else
				DisplayVoucher("1") 'Without Edit Option
				'document.formname.btnAdd.disabled = True
				'document.formname.btnVouDel.disabled = True
				document.formname.selAccHead.disabled = True
			End IF
			For iCtr = 0 To document.formname.selBook.length - 1
				IF Cstr(document.formname.selBook.Options(iCtr).Value) = Cstr(document.formname.hPreBookNo.value) Then
						document.formname.selBook.selectedIndex = iCtr
						Exit For
				End IF
			Next
		End IF
	End IF
End Function

Function DelVouch()
	document.formname.action = "VouDeletion.asp"
	document.formname.submit()
End Function


