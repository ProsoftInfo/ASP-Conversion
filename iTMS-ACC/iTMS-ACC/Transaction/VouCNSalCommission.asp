<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNSalCommission.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Ragavendran
	'Created On					:	Feb 16,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include file="../../include/populate.asp"-->
<%
Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1,sTxtNarration
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal,sBookName
dim sSalType,sOrgId,sOrgName,sQuery,sPartyName,sInvoiceNo,iInvNo
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff,sFromSal
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId,sTemp,iitemCode,iClassCode,iSalRetQty
Dim sRatePer,nArrCount,sTempVal,sCommtypename,sAgentCode,sSalTransNo

Dim sFinPeriod,sFinFrm,sFinTo,sValTemp,objDOM1,sdomRoot,sdomPage
Dim sCallFrom ,iCnt,sBookNumber,sInvFrom,sTempComm,sTempCommVal
Dim oDOMGJ,oDGjRoot,oDGjEntry,oDGjAcc,oDGjNarr,oDSubNode,nRatePer
sFinPeriod = Session("FinPeriod")
sValTemp = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp(0))
sFinTo = Trim(sValTemp(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"
iCnt =0
sBookNumber = Request.Form("selBook")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set oDOMGJ = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")
Set objDOM1 = Server.CreateObject("Microsoft.XMLDOM")
sInvoiceNo=Request.Form("selInvoiceNo")
Response.Write sInvoiceNo
'Response.end
sTemp = Split(sInvoiceNo,":")
sBookName=Request.Form("hBookName")

sCallFrom= Request.QueryString("hCallFrom")
sInvoiceNo = sTemp(0)
iInvNo = sInvoiceNo
sTempComm = Request.Form("hVouDetails")
sTempCommVal=split(sTempComm,",")
sUserid = getUserID()

sQuery = "Select FromApplication From Acc_T_CreatedVoucherHeader "&_
		 "Where CreatedTransNo = "&iInvNo&" and FromApplication is Not NULL "
Response.Write sQuery
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sFromSal = "Y"
Else
	sFromSal = "N"
End IF
objRs.Close
'For GJVoucher
Dim sDocNO,sInvNo,sInvDate,sTransAmt,sAmtAdjusted,sAmtToAdjust,sDocType
Dim sAmtToAcc,sPayableNo, sAdjType,sAccPay,sAccRec,sAccType,sAccAdv
Dim sRetVal
'oDOM.load server.MapPath("../xmldata/Voucher/"&sInvoiceNo&".xml")
sRetVal = GetVouchXML(sInvoiceNo)

oDOM.Load server.MapPath(sRetVal)
set oNodRoot = oDOM.documentElement
''CN Case XML Creation
for each oNodHeader in oNodRoot.childNodes
	if oNodHeader.nodeName="Header" then
		for Each oNodEntry in  oNodHeader.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sOrgName=oNodEntry.text
			end if
			if oNodEntry.nodeName="Book" then
				oNodEntry.Attributes.Item(0).nodeValue=Request.Form("selBook")
				oNodEntry.text=Request.Form("hBookName")
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sInvoiceNo=oNodEntry.Attributes.Item(0).nodeValue&"&nbsp;-&nbsp;"&oNodEntry.Attributes.Item(1).nodeValue
			end if
		next
	end if

	if oNodHeader.nodeName="Details" then
		set oNodDeatils=oNodHeader
	end if
	if oNodHeader.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodHeader
	end if
	if oNodHeader.nodeName="AgentDetails" then
		set oNodtemp=oNodRoot.removeChild(oNodHeader)
	end if
next

dim dInvAmount
sInvFrom = Request.Form("hInvVal")

'dInvAmount = sInvValue
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML id="TaxData" src="<%="../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml"%>"></XML>
<XML id="GJVoucher"></XML>
<SCRIPT LANGUAGE=javascript SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<script language="vbscript" >
'----------------------------------------------

'---------------------------------------------
Function SaveXML()

	Dim sExp,TempNode,sCheckVal,iCrAccHead,sDesc,sTempUnitNo,sTempUnitName
	Dim oDOMGJ,oDGjRoot,oDGjEntry,oDGjAcc,oDGjNarr,oDSubNode,sPartyName
	Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1
	Dim NodPayRec,NodDoc,NodRecCount,sPartyNo,sInvNoDet,sInvDate,sTempArr,sTempValue
	Dim sPartyDet,sArrValue,sRetVal2,setFlag
	set RootNode=TaxData.documentElement
	setFlag = false

	if trim(document.formname.SelAccountHd.value)="0" then
		alert("Select Account Head")
		exit function
	end if

	if trim(document.all.spAccHead.innerHTML)="" then
		alert("Select Account Head Value")
		exit function
	end if

if trim(document.formname.hCallFromVoucher.value)="CR" then

	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="Details" then
		 	oNodTemp.Attributes.GetNamedItem("VouDate").value=document.formname.ctlDate.GetDate
		end if
	next

	sExp = "//SaleInvoice"
	Set TempNode = RootNode.SelectNodes(sExp)

	Set newElem = TaxData.CreateElement("Narration")
	newElem.Text = document.formname.txtNarration.value
	RootNode.appendChild newElem


	IF document.formname.optApprove(0).checked = True Then
		sCheckVal = "Y"
		IF document.formname.selUserId.selectedIndex = 0 Then
			MsgBox "Select Approver "
			document.formname.selUserId.focus()
			Exit Function
		End IF
	Else
		sCheckVal = "N"
	End IF

	IF TempNode.length <> 0 Then
		Set newElem  = TaxData.createAttribute("Approval")
		newElem.value = sCheckVal
		TempNode.Item(0).setAttributeNode(newElem)

		Set newElem  = TaxData.createAttribute("Approver")
		newElem.value = document.formname.selUserId.value
		TempNode.Item(0).setAttributeNode(newElem)

		Set newElem  = TaxData.createAttribute("SalTrNo")
		newElem.value = document.formname.hdTransNo.Value
		TempNode.Item(0).setAttributeNode(newElem)

	End IF

	sExp = "//TaxDetails"
	Set TempNode = RootNode.SelectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.Item(0).nodeValue = document.formname.txtTotalInv.Value
	End IF

	IF document.formname.SelCrAgain.value = "A" Then
		Dim dEachItmVal
		sExp = "//Entry"
		Set TempNode = RootNode.SelectNodes(sExp)
		IF TempNode.length <> 0 Then
			For iCtr = 0 To TempNode.length - 1
				Set dEachItmVal = Eval("document.formname.txtAmount"&iCtr+1)
				TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value = dEachItmVal.Value
			Next
		End IF
	End IF

	IF document.formname.SelAccountHd.value = "G" Then
		iCrAccHead = document.formname.hCrAccHead.value
		sDesc = document.all.spAccHead.innerHTML
		sExp = "//AccHead"
		Set TempNode = RootNode.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			For iCtr = 0 To TempNode.length - 1
				TempNode.Item(iCtr).Attributes.getNamedItem("No").Value = iCrAccHead
				TempNode.Item(iCtr).Attributes.getNamedItem("Name").Value = sDesc
			Next
		End IF

		sExp = "//Tax"
		Set TempNode = RootNode.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			For iCtr = 0 To TempNode.length - 1
				TempNode.Item(iCtr).Attributes.getNamedItem("AccHead").Value = iCrAccHead
			Next
		End IF
	End IF


	IF CStr(document.formname.SelCrAgain.value) = "A" Then
		sExp = "//Tax"
		Set TempNode = RootNode.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			For iCtr = 0 To TempNode.length - 1
				TempNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value = "0.00"
			Next
		End IF
	End IF


	IF CheckFinDate Then
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","XMLSave.asp?Mod=CN&Name=Voucher Entry", false
		objhttp.send TaxData.XMLDocument
		if objhttp.responseText <> "" then
			Msgbox(objhttp.responseText)
		else
			if trim(document.formname.hInvCallFrom.value)="SI" then
				document.formname.action="VouCNOthInvGenerate.asp"
			elseif trim(document.formname.hInvCallFrom.value)="SR" then
				document.formname.action="VouCNSalRetAdj.asp"
			end if
			document.formname.submit()
		end if
	End IF
else 'if trim(document.formname.hCallFromVoucher.value)="CR" then

	set oNodRoot = TaxData.documentElement
	set oDGjRoot = GJVoucher.createElement("voucher")
	GJVoucher.appendChild oDGjRoot
	for each oNodHeader in oNodRoot.childNodes
		if oNodHeader.nodeName = "Header" then
			for each oDSubNode in oNodHeader.childNodes
				if oDSubNode.nodeName="Organization" then
				sTempUnitNo = oDSubNode.getAttribute("OrgId")
				sTempUnitName = oDSubNode.text
					oDGjRoot.setAttribute "UnitNo",oDSubNode.getAttribute("OrgId")
					oDGjRoot.setAttribute "UnitName",oDSubNode.text
					oDGjRoot.setAttribute "BookNo",document.formname.hBookCode.value
					oDGjRoot.setAttribute "BookName",document.formname.hBookName.value
					oDGjRoot.setAttribute "CRDR",""
				elseif oDSubNode.nodeName="Book" then
					sBookACHead = oDSubNode.getAttribute("BKAccHead")
				elseif oDSubNode.nodeName="Party" then
					sPartyNo = oDSubNode.getAttribute("ParType")&"?"&oDSubNode.getAttribute("ParSubType")&"?"&oDSubNode.getAttribute("ParType")&"-"& oDSubNode.getAttribute("ParSubTypeName")&"?"&oDSubNode.getAttribute("ParCode")
					sPartyNo = document.formname.hagentType.value&"?"&document.formname.hagentsubType.value&"?"&document.formname.hagentType.value&"- Commission Agent"&"?"&document.formname.hagentCode.value
				'	sPartyDet = sTempUnitNo &"&ParSubType="&oDSubNode.getAttribute("ParSubType")&"&ParType=" &oDSubNode.getAttribute("ParType")&"&PartyCode="&oDSubNode.getAttribute("ParCode")
					sPartyName = document.formname.hagentName.value
				elseif oDSubNode.nodeName="SaleInvoice" then
					sInvNoDet ="SALE INV NO:"&oDSubNode.getAttribute("InvNo") &" Dt:"& oDSubNode.getAttribute("InvDate")
					sInvDate = oDSubNode.getAttribute("InvDate")
				end if
			next
		elseif oNodHeader.nodeName="Details" then
			oDGjRoot.setAttribute "VouDate",document.formname.ctlDate.GetDate
			oDGjRoot.setAttribute "BookAcchead",sBookACHead
			oDGjRoot.setAttribute "Approver",document.formname.selUserId.value

			for each oNodEntry in oNodHeader.childNodes
				if oNodEntry.nodeName="Entry" then
					if setFlag = False then
						'First Entry
						set oDGjEntry= GJVoucher.createElement("Entry")
						iCnt = iCnt + 1
						oDGjEntry.setAttribute "No",iCnt
						oDGjEntry.setAttribute "CRDR","C"
						oDGjEntry.setAttribute "Payto","0"
						oDGjEntry.setAttribute "Amount",document.formname.txtTotalInv.Value
						oDGjEntry.setAttribute "AccUnit",sTempUnitNo
						oDGjEntry.setAttribute "AccName",sTempUnitName
						oDGjEntry.setAttribute "TdsAmount","0.00"
						oDGjEntry.setAttribute "TDSElgi","0"
						oDGjEntry.setAttribute "TdsPercentage","0"
						oDGjEntry.setAttribute "PayRecAmount","0"
						oDGjRoot.appendChild oDGjEntry
						for each oNodDeatils in oNodEntry.childNodes
							if oNodDeatils.nodeName="AccHead" then
								set oDGjAcc = GJVoucher.createElement("AccHead")
								oDGjAcc.setAttribute "No",sPartyNo
								oDGjAcc.setAttribute "Pay","0"
								oDGjAcc.setAttribute "Rec","0"
								oDGjAcc.setAttribute "Name",sPartyName
								oDGjAcc.setAttribute "Type","P"
								oDGjAcc.setAttribute "Adv","0"
								oDGjEntry.appendChild oDGjAcc
							end if

							Set oDGjNarr = GJVoucher.CreateElement("Narration")
							oDGjNarr.Text = document.formname.txtNarration.value
							oDGjEntry.appendChild oDGjNarr
						next
						setFlag = True
					end if ' 	if setFlag = False then

					'Second Entry
					set oDGjEntry= GJVoucher.createElement("Entry")
					iCnt = iCnt + 1
					oDGjEntry.setAttribute "No",iCnt
					oDGjEntry.setAttribute "CRDR","D"
					oDGjEntry.setAttribute "Payto",""
					oDGjEntry.setAttribute "Amount",document.formname.txtTotalInv.Value
					oDGjEntry.setAttribute "AccUnit",sTempUnitNo
					oDGjEntry.setAttribute "AccName",sTempUnitName
					oDGjEntry.setAttribute "TdsAmount","0.00"
					oDGjEntry.setAttribute "TDSElgi","0"
					oDGjEntry.setAttribute "TdsPercentage","0"
					oDGjEntry.setAttribute "PayRecAmount","0"
					oDGjRoot.appendChild oDGjEntry
					for each oNodDeatils in oNodEntry.childNodes
						if oNodDeatils.nodeName="AccHead" then
							set oDGjAcc = GJVoucher.createElement("AccHead")
							oDGjAcc.setAttribute "No",document.formname.hCrAccHead.value
							oDGjAcc.setAttribute "CostCenter",oNodDeatils.getAttribute("CostCenter")
							oDGjAcc.setAttribute "Analytical",oNodDeatils.getAttribute("Analytical")
							oDGjAcc.setAttribute "Name",document.all.spAccHead.innertext
							oDGjAcc.setAttribute "Type",oNodDeatils.getAttribute("Type")
							if oNodDeatils.getAttribute("No")<>"0" then
								oDGjAcc.setAttribute "TransFlag","A"
							else
								oDGjAcc.setAttribute "TransFlag","W"
							end if
							oDGjEntry.appendChild oDGjAcc
						end if
						Set oDGjNarr = GJVoucher.CreateElement("Narration")
						oDGjNarr.Text = document.formname.txtNarration.value
						oDGjEntry.appendChild oDGjNarr
					next
				end if
			next
		end if
	next


	set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","AccSalCommUpdate.asp", false
		objhttp.send

		IF CheckFinDate Then
			set objhttp = CreateObject("Microsoft.XMLHTTP")
			objhttp.Open "POST","XMLSave.asp?Mod=GJ&Name=Voucher Entry", false
			objhttp.send GJVoucher.XMLDocument
			if objhttp.responseText <> "" then
				Msgbox(objhttp.responseText)
			else
				document.formname.action="VouGenerate.asp"
				document.formname.submit()
			end if
		End IF
end if'if trim(document.formname.hCallFromVoucher.value)="CN" then
End Function

Function EnbApp(sObj)
	IF sObj.value = "Y" Then
		document.formname.selUserId.disabled = False
	Else
		document.formname.selUserId.selectedIndex = 0
		document.formname.selUserId.disabled = True
	End IF
End Function

Function SetRetVal(sObj,sCallType)

	Dim sVal,sPreInvVal,sNewInvVal,dDiffVal,sLen,sTotQty
	sVal = sObj.value
	sLen = document.formname.hRowVal.value

	document.formname.txtNarration.value = "CR Note for "& split(InoviceNo.innerText,"-")(0) &" "& split(InoviceNo.innerText,"-")(1) & " Sales Return Qty: "&	sTotQty

End Function

Function ResetTax()
	dim dInvAmount,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax
	set RootNode=TaxData.documentElement

	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next


	dInvAmount = 0
	dBasicTotal = 0
	dTotal = 0
	dInvAmount = CDbl(dInvAmount)
	dBasicTotal = CDbl(dBasicTotal)
	dTotal = CDbl(dTotal)

	For Each oNodEntry in TaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue
		oNodEntry.Attributes.Item(5).nodeValue="0.00"
		eval("document.formname.txtTaxValue"&sCatCode&sTaxCode).value="0.00"
	next

	dInvAmount=FormatNumber(dInvAmount,2,,,0)
	document.formname.txtInvValue.value=dInvAmount
	Taxroot.Attributes.Item(0).Nodevalue = "0.00"
	TaxRoot.Attributes.Item(1).nodeValue="0.00"
	TaxRoot.Attributes.Item(2).nodeValue="0.00"

End Function

Function CheckNoSer()
	Dim ObjHttp,sPassVal,sMon,sYear,sDate,sPeriod,sRetVal
	Set ObjHttp = CreateObject("MSXML2.XMLHTTP")
	IF Cstr(document.formname.hVouCode.Value) = "04" Then
		sPassVal = document.formname.selUnitId.Value 'Unit
		sPassVal = sPassVal&":"&document.formname.hVouCode.Value 'BookCode
		sPassVal = sPassVal&":"&document.formname.hCallFrm.Value 'Call Fro Created or Accounted
		sPassVal = sPassVal&":D"
		sPassVal = sPassVal&":"&document.formname.selBook.Value 'Book Number
	Else
		sPassVal = document.formname.hOrgid.Value 'Unit
		sPassVal = sPassVal&":"&document.formname.hVouCode.Value 'BookCode
		sPassVal = sPassVal&":"&document.formname.hCallFrm.Value 'Call Fro Created or Accounted
		IF Cstr(document.formname.hVouCRDR.Value) = "" Then
			sPassVal = sPassVal&":"&"D"
		Else
			sPassVal = sPassVal&":"&document.formname.hVouCRDR.Value ' Voucher Type C/D
		End IF
		sPassVal = sPassVal&":"&document.formname.hBookCode.Value 'Book Number
	End IF

	sPeriod = document.formname.ctlDate.GetDate()
	sPassVal = sPassVal&":"&sPeriod 'Voucher Date
	ObjHttp.open "GET","NoSeriesCheck.asp?sValue="&sPassVal, False
	ObjHttp.send
	sRetVal = ObjHttp.responseText
	IF CStr(Trim(sRetVal)) = "T" Then
		CheckNoSer = True
	Elseif CStr(Trim(sRetVal)) = "F" Then
		MsgBox "No Series is Not Defined "
		CheckNoSer = False
		Exit Function
	Else
		MsgBox "Error "
		CheckNoSer = False
		Exit Function
	End IF
End Function

Function AccHead(objAcc)
dim sOrgId,sTemp,sDesc
	sOrgId=document.formname.hOrgId.value
	if objAcc.selectedIndex >0 then
		if objAcc.value="G" then
			showGLHead sOrgId
		End if 'End of select Account Head Type check GL or PARTY
	End if 'End of If any Account Head Selected Check
End function

Function showGLHead(sOrgId)
	Dim iAccCode,sRetVal,arrTemp,sDesc,sTemp

	iBookNo=document.formname.hBookcode.value

	OutValue = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")

	while UBound(arrTemp) = 0
		OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend

	sRetVal = OutValue

	if UBound(arrTemp) <= 1 then exit function
	sTemp = Split(sRetVal,":")
	iAccCode = sTemp(0)
	sDesc = sTemp(5)

	document.formname.hCrAccHead.value = iAccCode
	document.all.spAccHead.innerHTML = sDesc


End function

Function ReTotalCr()
	Dim sExp,TempNode,RootNode,iTax,iCat,dTaxVal,iCtr,sObj,dSalVal,dCrVal
	dSalVal = document.formname.txtTotal.Value
	set RootNode=TaxData.documentElement
	sExp = "//Tax"
	Set TempNode = RootNode.selectNodes(sExp)
	dTaxVal = 0
	IF TempNode.Length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iTax = TempNode.Item(iCtr).Attributes.getNamedItem("TaxCode").Value
			iCat = TempNode.Item(iCtr).Attributes.getNamedItem("CatCode").Value
			Set sObj = Eval("document.formname.txtTaxValue"&iCat&iTax)
			TempNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value = sObj.Value
			dTaxVal = Cdbl(dTaxVal) + CDbl(sObj.Value)
		Next
	End IF
	dCrVal = CDbl(dSalVal) + CDbl(dTaxVal)
	dCrVal = FormatNumber(dCrVal,2,,,0)
	document.formname.txtInvValue.Value = dCrVal
	document.formname.txtTotalInv.value = dCrVal

End Function

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

	'MsgBox sCurrMonYr
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
'==================================
Function init()
	document.all.sPartyName.innertext = document.formname.hAgentName.value
	document.formname.txtNarration.value =  document.formname.hTxtnarr.value
End Function
'==========================================================================================================================


</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="init()">
<form method="POST" name="formname">
<Input type="hidden" name="hInvCallFrom" value="<%=sInvFrom%>">
<Input type="hidden" name="hdTransNo" value="<%=iInvNo%>">
<Input type="hidden" name="hAccType" value="C">
<Input type="hidden" name="hCallType" value="OINV">
<Input type="hidden" name="hNoteType" value="C">
<Input type="hidden" name="hFromSal" value="<%=sFromSal%>">
<Input type="hidden" name="hOrgid" value="<%=sOrgId%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<Input type="hidden" name="hBookCode" value="<%=Request.Form("selBook")%>">
<input type="hidden" name="hCrAccHead" value="0">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">
<input type="hidden" name="hCallFromVoucher" value="<%=sCallFrom%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<%IF sCallFrom= "CR" then%>
<input type="hidden" name="hVouCode" value="07">
<%else%>
<input type="hidden" name="hVouCode" value="08">
<%end if%>
<input type="hidden" name="hVouType" value="">
<input type="hidden" name="hVouName" value="GJ">
<input type="hidden" name="optApprove" value="">
<input type="hidden" name="hInsVou" value="">
<input type="hidden" name="hSelVouDate" value="">
<input type="hidden" name="hTransNo" value="<%=iInvNo%>">
<input type="hidden" name="CallType" value="CN">
<input type="hidden" name="CrVouNo" value="">
<input type="hidden" name="SelTDSGrp" value="">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		Credit Note For Sales Commission

		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<tr><td align="center">Voucher</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <!--tr>
                            <td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr-->
                            <!--tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: hand" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">?</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Daywise Balance"><font size="3" face="Webdings">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Voucher History">
                    <font size="4" face="Webdings">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell">
                    &nbsp;
                    </td>
                        </tr>
                            </table>
                            </td>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0" width="575">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Unit </td>
                                                    <td class="FieldCell">  <span class="DataOnly"><%=sOrgName%></span></td>
                                                    <td class="FieldCellSub" width="75"><p align="left">Date</p></td>
                                                    <td class="FieldCellSub" width="145">
                                                 <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>
													</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Party
                                                      Name</td>
                                                    <td class="FieldCell" colspan="3">  <span id="sPartyName" class="DataOnly">&nbsp;</span></td>

                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Cr Note Against</td>
                                                    <td class="FieldCell" width="200"><span class="DataOnly">Sales Commission</span></td>
                                                    </tr>
                                                    <tr>
														<td class="FieldCellSub" width="200">Select Account Head</td>
														<td class="FieldCell" width="200">
														<Select name="SelAccountHd" class="FormElem" onChange="AccHead(this)">
															<Option Value="0">Select</Option>
															<Option Value="G">GL ACCOUNT HEAD</Option>
														</Select>
														</td>
														 <td class="FieldCellSub" colspan="2"><span class="DataOnly" id="spAccHead"></span>
														 </td>
                                                     </tr>


                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                            <tr>
                <td></td>
                <td valign="top" width="100%">
                <div class="frmBody" id="frm2" style="width: 575; height:150;">
            <table border="0" cellspacing="1" class="ExcelTable" width="575">
			    <tr>
					<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
					<td class="ExcelHeaderCell" align="center" colspan="2">Invoice</td>
					<td class="ExcelHeaderCell" align="center"  rowspan="2" width="150">Commission Type</td>
					<td class="ExcelHeaderCell" align="center"  rowspan="2" width="150">Commission Value</td>
			    </tr>
			    <tr>
						<td class="ExcelHeaderCell" align="center"  width="150" >Number</td>
						<td class="ExcelHeaderCell" align="center"  width="150" >Date</td>
			    </tr>
						<%

							set sdomRoot=objDOM1.createElement("Root")
							objdom1.appendChild sdomRoot

							iSno = 0
							sTxtNarration = "CR Note for "
							for nArrCount = 0 to UBound(sTempCommVal) - 1
								sTempVal = split(sTempCommVal(nArrCount),":")
								iSno = iSno + 1

								if Trim(sTempVal(5)) = "A" then
									sCommtypename = "Per Qty"
								elseif Trim(sTempVal(5)) = "B" then
									sCommtypename = "% of Basic"
								elseif Trim(sTempVal(5)) = "C" then
									sCommtypename = "% of Total Inv"
								elseif Trim(sTempVal(5)) = "P" then
									sCommtypename = "Per Packing"
								End if

								sTxtNarration = sTxtNarration & " : "& sTempVal(1) &" - "& sTempVal(2)
						%>
						    <tr>
						    <td class="ExcelSerial" align="center"><%=isno%></td>
						    <td class="ExcelDisplayCell"><%=sTempVal(1)%></td>
						    <td class="ExcelDisplayCell" ><%=sTempVal(2)%></td>
						    <td class="ExcelDisplayCell" align="Center"><%=sCommtypename%></td>
						    <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sTempVal(3),2,,,-1)%></td>
						    </tr>
						<%
								sAgentCode = sTempVal(8)
								dTotal = cdbl(dTotal) + cdbl(sTempVal(3))

								set sdomPage=objDOM1.createElement("CommDet")
								sdomPage.setAttribute "AgentCode",sTempVal(8)
								sdomPage.setAttribute "CommissionType",sTempVal(5)
								sdomPage.setAttribute "CurrCode",sTempVal(9)
								sdomPage.setAttribute "CommValue",sTempVal(3)
								objRs.Open "Select Saletransactionno from VwSalCommAccDet where TransactionNumber = "& sTempVal(0),con
								if not objRs.EOF then
									sSalTransNo = objRs(0)
								end if
								objRs.Close
								sdomPage.setAttribute "SalTransNo",sSalTransNo
								sdomPage.setAttribute "AccTransNo",sTempVal(0)
								sdomRoot.appendChild sdomPage

							next

							objDOM1.save server.MapPath("../temp/transaction/VoucherSalCommDet_"&Session.SessionID&".xml")

							with objRs
								.CursorLocation = 3
								.CursorType = 3
								.ActiveConnection = con
								.Source = "Select PartyType,PartySubType,R.PartyCode,PartyName from App_M_PartyMaster P,APP_R_orgparty R where  R.PartyCode = P.PartyCode and R.PartyCode = "& sAgentCode
								'Response.Write objRs.Source
								.Open
							end with
							if not objRs.EOF then
							%>
								<input type=hidden name="hAgentType" value="<%=objrs(0)%>">
								<input type=hidden name="hAgentSubType" value="<%=objrs(1)%>">
								<input type=hidden name="hAgentCode" value="<%=objrs(2)%>">
								<input type=hidden name="hAgentName" value="<%=objrs(3)%>">
							<%end if%>
        <tr>
			<Input type="hidden" name="hRowVal" value="<%=isno%>">
			<td class="ExcelSerial" align="center" colspan=4 ><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
			<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTotalInv" value="<%=FormatNumber(dTotal,2,,,0)%>" class="Formelem" size="13"></td>
			 <td align="right" colspan="3"></td>
        </tr><input type="hidden" name="hTxtnarr" value="<%=sTxtNarration%>">
			<%
				oDOM.save server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")
			%>
	        </table>

                </div>
                </td>
                <td></td>
                            </tr>
                            <tr>
                            <td></td>
			<td align="left" class="FieldCellSub"  valign="Top">
				Approval &nbsp;&nbsp;&nbsp;

			<Input type="radio" name="optApprove" checked value="Y" onClick="EnbApp(this)"> Yes &nbsp;&nbsp;&nbsp;
			<Input type="radio" name="optApprove" value="N" onClick="EnbApp(this)"> No &nbsp;&nbsp;&nbsp;
			</td>
        </tr>
        <tr>
			 <td></td>
			<td align="left" class="FieldCellSub"  valign="Top">
				Immediate Approver &nbsp;&nbsp;&nbsp;

			<select size="1" name="selUserId" class="FormElem">
              <option value="I">Immediate Approver</option>
                <%=populateEmployeeWithVal(sUserId)%>
                    </select>
			</td>
        </tr>
        <tr>
			 <td></td>
			<td align="left" class="FieldCellSub"  valign="Top">
				Narration &nbsp;&nbsp;&nbsp;

			<Textarea name="txtNarration" class="FormElem" cols="40" rows="4" maxlength="200" readonly ></Textarea>

			</td>
        </tr>

                                                        <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                                        </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
                                                                <input type="button" value="Next" name="B2" class="ActionButton" onClick="SaveXML()" >
                                                                <input type="button" value="Cancel" name="B6" class="ActionButton" onClick="Cancel('VOUCNBOOKSELECTION.ASP')" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="BottomPack" colspan="3">
								</td>
                                </tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</html>
