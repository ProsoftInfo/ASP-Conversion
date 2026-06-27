<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNSalInvAmd.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	October 20, 2004
	'Modified On				:
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
Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal,sBookName
dim sSalType,sOrgId,sOrgName,sQuery,sPartyName,sInvoiceNo,iInvNo
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff,sFromPur
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId,sTemp
Dim sFromApp,dOrgQty,dOrgRate,dOrgDis,iTransNo,sCrAgain,dTotQty,dTotRate,dTotDis
Dim ObjInv,sNarration,iCrAltAccHd,sCrAltAccName,sAccChg,sAmdType,iItemCode,iClassCode,dInvInvAmount

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set ObjInv = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

iTransNo=Request.Form("hTransNo")
sBookName=Request.Form("hBookName")
sFromApp = Request.Form("hFromApp")

sTemp = Split(iTransNo,"-")
iTransNo = sTemp(0)
sAmdType = Request("AmdType")

sUserid = getUserID()

Dim sRetVal
'oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)


'Response.Write iTransNo

set oNodRoot=oDOM.documentElement

for each oNodHeader in oNodRoot.childNodes
	if oNodHeader.nodeName="Header" then
		for Each oNodEntry in  oNodHeader.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sOrgName=oNodEntry.text
			end if
			'if oNodEntry.nodeName="Book" then
			'	oNodEntry.Attributes.Item(0).nodeValue=Request.Form("selBook")
			'	oNodEntry.text=Request.Form("hBookName")
			'end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sInvoiceNo=oNodEntry.Attributes.Item(0).nodeValue&"&nbsp;-&nbsp;"&oNodEntry.Attributes.Item(1).nodeValue
				iInvNo = oNodEntry.Attributes.getNamedItem("SalTrNo").Value
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
dim dInvAmount,sExp,TempNode,iCtr

dTotQty = 0
dTotRate = 0
dTotDis = 0

sQuery = "Select Sum(InvoicedQuantity),Sum(InvoicedRate),Sum(DiscountPercent) From  "&_
		 "Acc_T_CreatedVoucherDetails Where CreatedTransNo = "&iInvNo&" "

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	dOrgQty = objRs(0)
	dOrgRate = objRs(1)
	dOrgDis = objRs(2)
End IF
objRs.Close

sQuery = "Select isNull(PurchaseBillType,'') From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	sCrAgain = objRs(0)
End IF
objRs.Close


sExp = "//Entry"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.Length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		dTotQty = CDbl(dTotQty) + CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("Qty").Value)
		dTotRate = CDbl(dTotRate) + CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("Rate").Value)
		dTotDis = CDbl(dTotDis) + CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("DisPer").Value)
	Next
End IF

'IF CDbl(dOrgQty) <> CDbl(dTotQty) Then
'	sCrAgain = "Q"
'Elseif CDbl(dOrgRate) <> CDbl(dTotRate) Then
'	sCrAgain = "R"
'Elseif CDbl(dOrgDis) <> CDbl(dTotDis) Then
'	sCrAgain = "D"
'Else
'	sCrAgain = "A"
'End IF


sInvValue=oNodTaxRoot.Attributes.Item(2).nodeValue
dInvAmount=oNodTaxRoot.Attributes.Item(0).nodeValue




sExp = "//Voucher/Narration"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.Length <> 0 Then
	sNarration = Trim(TempNode.Item(0).Text)
End IF

sExp = "//AccHead"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	iCrAltAccHd = TempNode.Item(0).Attributes.getNamedItem("No").Value
	sCrAltAccName = TempNode.Item(0).Attributes.getNamedItem("Name").Value
End IF

sQuery = "Select AccUnitAccountHead From Acc_T_CreatedVoucherDetails Where  "&_
		 "AccUnitAccountHead = "&iCrAltAccHd&" and CreatedTransNo = "&iInvNo&"  "&_
		 "and VoucherEntryNumber = 1 "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sAccChg = "N"
	iCrAltAccHd = 0
Else
	sAccChg = "Y"
End IF
objRs.Close

Response.Write iInvNo &"==============="

sRetVal = GetVouchXML(iInvNo)
ObjInv.Load server.MapPath(sRetVal)
Dim oInvNodRoot,oInvNodHeader,oInvNodEntry,oInvNodTaxRoot,oInvNodDeatils
set oInvNodRoot=ObjInv.documentElement

for each oInvNodHeader in oInvNodRoot.childNodes
	if oInvNodHeader.nodeName="Details" then
		set oInvNodDeatils=oInvNodHeader
	end if
	if oInvNodHeader.nodeName="TaxDetails" then
		set oInvNodTaxRoot=oInvNodHeader
	end if
next
dInvInvAmount=oInvNodTaxRoot.Attributes.Item(0).nodeValue
ObjInv.save server.MapPath("../temp/transaction/InvDet_ForDN_"&Session.SessionID&".xml")



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<XML id="TaxData" src="<%="../temp/transaction/Voucher EntryAmd_DN_"&Session.SessionID&".xml"%>"></XML>
<XML id="InvData" src="<%="../temp/transaction/InvDet_ForDN_"&Session.SessionID&".xml"%>"></XML>

<SCRIPT LANGUAGE=javascript SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<script language="vbscript" >
Dim InvRoot
Set InvRoot = InvData.documentElement


Function SaveXML()

	Dim sExp,TempNode,sCheckVal,iCtr,iCrAccHead,sDesc,TempInvNode
	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="Details" then
		 	oNodTemp.Attributes.GetNamedItem("VouDate").value=document.formname.ctlDate.GetDate
		end if
	next



	sExp = "//Voucher/Narration"
	Set TempNode = RootNode.SelectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Text = document.formname.txtNarration.value
	Else
		Set newElem = TaxData.CreateElement("Narration")
		newElem.Text = document.formname.txtNarration.value
		RootNode.appendChild newElem
	End IF

	sExp = "//SaleInvoice"
	Set TempNode = RootNode.SelectNodes(sExp)




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
		newElem.value = document.formname.hdTransNo.value
		TempNode.Item(0).setAttributeNode(newElem)

	End IF

	IF document.formname.SelCrAgain.value = "A" Then
		sExp = "//TaxDetails"
		Set TempNode = RootNode.SelectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.Item(0).nodeValue = document.formname.txtCrNoteValue.value
		End IF
	End IF
''blocked by ragav on Jan 25,2012
''Begin
'	IF document.formname.SelCrAgain.value = "A" Then
'		sExp = "//Entry"
'		Set TempNode = RootNode.SelectNodes(sExp)
'		IF TempNode.length <> 0 Then
'			For iCtr = 0 To TempNode.length - 1
'				TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value = document.formname.txtCrNoteValue.value
'			Next
'		End IF
'	End IF
''End
	'MsgBox document.formname.SelAccountHd.value
	Set InvRoot = InvData.documentElement
	'MsgBox InvRoot.xml

	Dim iEntryNo,iEntAccNo,sEntAccDesc
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



'	IF CStr(document.formname.SelCrAgain.value) = "A" Then
'		sExp = "//Tax"
'		Set TempNode = RootNode.selectNodes(sExp)
'		IF TempNode.length <> 0 Then
'			For iCtr = 0 To TempNode.length - 1
'				TempNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value = "0.00"
'			Next
'		End IF
'	End IF

	'MsgBox RootNode.xml

	'IF CheckNoSer() Then
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","XMLSave.asp?Mod=DN&Name=Voucher AMD", false
		objhttp.send TaxData.XMLDocument
		if objhttp.responseText <> "" then
			Msgbox(objhttp.responseText)
		else
			'alert(TaxData.xml)
			'MsgBox "OK "
			document.formname.submit()
		end if
	'End IF
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
	Dim sVal,sPreInvVal,sNewInvVal,dDiffVal,dRowLen,InvRoot,iSno
	Dim sExp,TempNode,iCtr
	sVal = sObj.value
	dRowLen = document.formname.hRowVal.value
	Set InvRoot = InvData.documentElement

	IF sCallType = "1" Then
		document.formname.reset()
		For iCtr = 0 To sObj.length - 1
			IF sObj.Options(iCtr).value = sVal Then
				sObj.selectedIndex = iCtr
				Exit For
			End IF
		Next
	ElseIF sCallType = "2" Then
		ResetCrVal()
	End IF

	dPreInvVal = document.all.txtFInvValue.value
	dPreInvVal = CDbl(dPreInvVal)
	sNewInvVal = CDbl(document.formname.txtInvValue.Value)

	'IF CStr(document.formname.hAccType.Value) = "C" Then
	'	dDiffVal = CDbl(dPreInvVal) - CDbl(sNewInvVal)
	'Else
	'	dDiffVal = CDbl(sNewInvVal) - CDbl(dPreInvVal)
	'End IF

	'dDiffVal = FormatNumber(dDiffVal,2,,,0)
	sNewInvVal = FormatNumber(sNewInvVal,2,,,0)

	For iCtr = 1 to dRowLen
		IF CStr(sVal) = "Q" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElem"
			Eval("document.all.tQty"&iCtr).className = "ExcelInputCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"

			document.formname.txtCrNoteValue.value = sNewInvVal

			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = False
			Eval("document.formname.txtRate"&iCtr).readonly = True
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tDrVal.classname = "ExcelDIsplayCell"
		End IF

		IF CStr(sVal) = "R" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElem"
			Eval("document.all.tRate"&iCtr).className = "ExcelInputCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"

			document.formname.txtCrNoteValue.value = sNewInvVal

			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = False
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tDrVal.classname = "ExcelDIsplayCell"
		End IF

		IF CStr(sVal) = "D" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElem"
			Eval("document.all.tDis"&iCtr).className = "ExcelInputCell"

			document.formname.txtCrNoteValue.value = sNewInvVal

			Eval("document.formname.txtdis"&iCtr).readonly = False
			Eval("document.formname.txtQty"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = True
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tDrVal.classname = "ExcelDIsplayCell"
		End IF

		IF CStr(sVal) = "0" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"

			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = True
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tDrVal.classname = "ExcelDIsplayCell"

		End IF


		IF CStr(sVal) = "A" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			'Eval("document.all.tDis"&iCtr).class = "ExcelDisplayCell"

			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = True
			document.formname.txtCrNoteValue.readOnly = False
			document.formname.txtCrNoteValue.className = "FormElem"
			document.all.tDrVal.classname = "ExcelInputCell"

			ResetTax()
		End IF

	Next
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
	'sMon = Mid(sDate,4,2)
	'sYear = Right(sDate,4)
	'sPeriod = Trim(sYear)&Trim(sMon)
	sPassVal = sPassVal&":"&sPeriod 'Voucher Date

	'MsgBox sPassVal

	ObjHttp.open "GET","NoSeriesCheck.asp?sValue="&sPassVal, False
	ObjHttp.send
	sRetVal = ObjHttp.responseText

	'alert sRetVal

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

Function ResetInvVal()
	Dim InvRoot,sExp,TempNode,iCtr,iSno,iTmp1,iTmp2
	Set InvRoot = InvData.documentElement
	sExp = "//Entry"
	Set TempNode = InvRoot.selectNodes(sExp)
	SetRetVal document.formname.SelCrAgain,1


	IF TempNode.Length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iSno = CDbl(iCtr) + 1
			Eval("document.all.tOldQty"&iSno).innerHTML = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("Qty").Value,2,,,0)
			Eval("document.all.tOldRate"&iSno).innerHTML = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("Rate").Value,2,,,0)
			Eval("document.all.tOldDis"&iSno).innerHTML = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("DisPer").Value,2,,,0)
			Eval("document.all.txtOldAmount"&iSno).Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value,2,,,0)
		Next
	End IF

	sExp = "//Details"
	Set TempNode = InvRoot.selectNodes(sExp)
	IF TempNode.Length <> 0 Then
		document.formname.txtTotalInv.Value = FormatNumber(TempNode.Item(0).Attributes.getNamedItem("BasicValue").Value,2,,,0)
		document.formname.txtFInvValue.Value = FormatNumber(TempNode.Item(0).Attributes.getNamedItem("BasicValue").Value,2,,,0)
	End IF

	sExp = "//TaxDetails"
	Set TempNode = InvRoot.selectNodes(sExp)
	IF TempNode.Length <> 0 Then
		document.formname.txtFInvValue.Value = FormatNumber(TempNode.Item(0).Attributes.getNamedItem("InvoiceVlaue").Value,2,,,0)
	End IF




	sExp = "//Tax"
	Set TempNode = InvRoot.selectNodes(sExp)
	IF TempNode.Length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iTmp1 = TempNode.Item(iCtr).Attributes.getNamedItem("CatCode").Value
			iTmp2 = TempNode.Item(iCtr).Attributes.getNamedItem("TaxCode").Value
			Eval("document.all.txtOldTaxValue"&iTmp1&iTmp2).Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value,2,,,0)
		Next
	End IF

End Function

Function ResetCrVal()
	Dim InvRoot,sExp,TempNode,iCtr,iSno,iTmp1,iTmp2,TempTaxNode
	Set InvRoot = InvData.documentElement
	set RootNode = TaxData.documentElement

	sExp = "//Entry"
	Set TempNode = InvRoot.selectNodes(sExp)
	SetRetVal document.formname.SelCrAgain,1

	IF TempNode.Length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iSno = CDbl(iCtr) + 1
			Eval("document.all.txtqty"&iSno).Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("Qty").Value,2,,,0)
			Eval("document.all.txtRate"&iSno).Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("Rate").Value,2,,,0)
			Eval("document.all.txtDis"&iSno).Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("DisAmount").Value,2,,,0)
			Eval("document.all.txtAmount"&iSno).Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value,2,,,0)
			sExp = "//Entry[@No="&iSno&"]"
			Set TempTaxNode = RootNode.selectNodes(sExp)
			IF TempTaxNode.length <> 0 Then
				TempTaxNode.Item(0).Attributes.getNamedItem("Qty").Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("Qty").Value,2,,,0)
				TempTaxNode.Item(0).Attributes.getNamedItem("Rate").Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("Rate").Value,2,,,0)
				TempTaxNode.Item(0).Attributes.getNamedItem("DisPer").Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("DisPer").Value,2,,,0)
				TempTaxNode.Item(0).Attributes.getNamedItem("Amount").Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value,2,,,0)
			End IF
		Next
	End IF


	sExp = "//Details"
	Set TempNode = InvRoot.selectNodes(sExp)
	IF TempNode.Length <> 0 Then
		document.formname.txtTotal.Value = FormatNumber(TempNode.Item(0).Attributes.getNamedItem("BasicValue").Value,2,,,0)

	End IF

	sExp = "//TaxDetails"
	Set TempNode = InvRoot.selectNodes(sExp)
	IF TempNode.Length <> 0 Then
		document.formname.txtInvValue.Value = FormatNumber(TempNode.Item(0).Attributes.getNamedItem("InvoiceVlaue").Value,2,,,0)
	End IF

	sExp = "//Tax"
	Set TempNode = InvRoot.selectNodes(sExp)
	IF TempNode.Length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iTmp1 = TempNode.Item(iCtr).Attributes.getNamedItem("CatCode").Value
			iTmp2 = TempNode.Item(iCtr).Attributes.getNamedItem("TaxCode").Value
			'MsgBox iTmp1 &"  " & iTmp2
			Eval("document.all.txtTaxValue"&iTmp1&iTmp2).Value = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value,2,,,0)
		Next
	End IF

End Function

Function AccHead(objAcc)
dim sOrgId,sTemp,sDesc
	sOrgId=document.formname.hOrgId.value
	'MsgBox objAcc.value
	if objAcc.value="G" then
		showGLHead sOrgId
	End if 'End of select Account Head Type check GL or PARTY
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
	document.formname.txtCrNoteValue.value = dCrVal

End Function
'************************
Function Init(sCrAgain)
set obj = eval("document.formname.SelCrAgain")
SetRetVal obj,1
popTax
End Function
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init('<%=sCrAgain%>')">
<%IF CStr(sAmdType) = "A" Then %>
	<form method="POST" name="formname" action="AmdAccCrNtGenerate.asp">
<%Else%>
	<form method="POST" name="formname" action="VouCNSalInvAmdUpdate.asp">
<%End IF%>
<Input type="hidden" name="hdTransNo" value="<%=iInvNo%>">
<Input type="hidden" name="hAccType" value="C">
<Input type="hidden" name="hCallType" value="OINV">
<Input type="hidden" name="hNoteType" value="C">
<Input type="hidden" name="hFromPur" value="<%=sFromPur%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hVouCode" value="07">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="<%=Request.Form("selBook")%>">
<input type="hidden" name="hCrTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hCrAccHead" value="<%=iCrAltAccHd%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		Credit Note For Sales Invoice Amendment
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
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
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0" width="100%">
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
                                                    <td class="FieldCell" colspan="3">  <span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>

                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Invoice No-Date</td>
                                                    <td class="FieldCell" width="200">  <span class="DataOnly"><%=sInvoiceNo%></span></td>
                                                    <td class="FieldCellSub" width="100">Invoice Value</td>
                                                    <td class="FieldCellSub" width="145"> <span class="DataOnly"><%=FormatNumber(sInvValue,2,,,0)%> </span></td>
                                                        </tr>

                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Cr Note Against</td>
                                                    <td class="FieldCell" width="200">
                                                    <Select name="SelCrAgain" class="FormElem" onChange="SetRetVal(this,'2')">
                                                    <Option Value="0">Select</Option>
                                                    <%IF CStr(sCrAgain) = "Q" Then %>
														<Option Value="Q" Selected>Quantity</Option>
													<%Else%>
														<Option Value="Q">Quantity</Option>
													<%End IF
													  IF CStr(sCrAgain) = "R" Then
													%>
														<Option Value="R" Selected>Rate</Option>
                                                    <%Else%>
														<Option Value="R">Rate</Option>
													<%End IF
													  IF CStr(sCrAgain) = "D" Then	%>
														<Option Value="D" Selected>Discount</Option>
													<%Else%>
														<Option Value="D">Discount</Option>
													<%End IF
													  IF CStr(sCrAgain) = "A" Then
													%>
														<Option Value="A" Selected>Quality</Option>
                                                    <%Else%>
														<Option Value="A">Quality</Option>
													<%End IF %>
                                                    </Select>
                                                    </td>
                                                    </tr>
                                                    <tr>
														<td class="FieldCellSub" width="200">Select Account Head</td>
														<td class="FieldCell" width="200">
														<Select name="SelAccountHd" class="FormElem" onChange="AccHead(this)">
														<%IF CStr(sAccChg) = "N" Then %>
																<Option Value="0" Selected>ITEM ACCOUNT HEAD</Option>
																<Option Value="G">GL ACCOUNT HEAD</Option>
														<%Else%>
																<Option Value="G"  Selected>GL ACCOUNT HEAD</Option>
														<%End IF %>
														</Select>
														&nbsp; <a href="javascript:AccHead(document.formname.SelAccountHd)"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Account Head"></a>
														</td>
														<td class="FieldCellSub" colspan="2">
														<%IF CStr(sAccChg) = "N" Then %>
															<span class="DataOnly" id="spAccHead"></span>
														<%Else%>
															<span class="DataOnly" id="spAccHead"><%=sCrAltAccName%></span>
														<%End IF%>
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
                <div class="frmBody" id="frm2" style="width: 775; height:245;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
    <td class="ExcelHeaderCell" align="center" colspan="4">Invoice</td>
    <td class="ExcelHeaderCell" align="center" colspan="4">Returned</td>
     <!--td class="ExcelHeaderCell" align="center" colspan="2">Invoice Value</td-->

        </tr>
        <tr>
    <td class="ExcelHeaderCell" align="center" width="55">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="55">Rate</td>
    <td class="ExcelHeaderCell" align="center" width="55">Discount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Amount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="55">Rate</td>
    <td class="ExcelHeaderCell" align="center" width="55">Discount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Amount</td>
        </tr>
<%
    Dim iInvSno,sInvDescription,sInvAmount,sInvQty
    Dim sInvInvValue,sInvRate,sInvRatePer,sInvDiscount,sInvDiscPer,dInvTotal,iInvItemCode,iInvClassCode
	Dim sRatePer
	
	For Each oInvNodEntry in oInvNodDeatils.childNodes
	    iInvSno=oInvNodEntry.getAttribute("No")
		sInvDescription=oInvNodEntry.getAttribute("PayTo")
		sInvAmount=oInvNodEntry.getAttribute("Amount")
		sInvQty=oInvNodEntry.getAttribute("Qty")
		sInvInvValue=oInvNodEntry.getAttribute("ActValue")
		sInvRate = oInvNodEntry.getAttribute("Rate")
		sInvRatePer = 1
		sInvDiscount=oInvNodEntry.getAttribute("DisAmount")
		sInvDiscPer = oInvNodEntry.getAttribute("DisPer")
		iInvItemCode = oInvNodEntry.getAttribute("ItemCode")
		iInvClassCode = oInvNodEntry.getAttribute("ClassCode")
		dInvTotal=CDbl(dInvTotal)+CDbl(sInvAmount)
		
		
		sCheckExp = "//Entry[@ItemCode="& iInvItemCode &" and @ClassCode="& iInvClassCode &" and @Amount>0]"
		set oNodtemp = oNodDeatils.selectNodes(sCheckExp)
		Response.Write "<font color=red>"
		if oNodtemp.length> 0 then
            sAmount=oNodtemp.Item(0).Attributes.GetNamedItem("Amount").value
            sQty=oNodtemp.Item(0).Attributes.GetNamedItem("Qty").value
            sValue=oNodtemp.Item(0).Attributes.GetNamedItem("ActValue").value
            sRate = oNodtemp.Item(0).Attributes.GetNamedItem("Rate").value
            sRatePer = 1
            sDiscount=oNodtemp.Item(0).Attributes.GetNamedItem("DisAmount").value
            sDiscPer = oNodtemp.Item(0).Attributes.GetNamedItem("DisPer").Value
            iItemCode = oNodtemp.Item(0).getAttribute("ItemCode")
            iClassCode = oNodtemp.Item(0).getAttribute("ClassCode")
            dTotal=CDbl(dTotal)+CDbl(sAmount)
        else
            sAmount="0"
            sValue="0"
            sQty = "0"
            sRate = "0"
            sRatePer = 1
            sDiscount="0"
            sDiscPer = "0"
            iItemCode = "0"
            iClassCode = "0"
            if sCrAgain="Q" then
                sQty="0"
                sRate = sInvRate 
                sDiscount = sInvDiscount 
                sDiscPer = sInvDiscPer 
            elseif sCrAgain = "R" then
                sQty=sInvQty 
                sRate = "0"
                sDiscount = sInvDiscount 
                sDiscPer = sInvDiscPer 
            elseif sCrAgain = "D" then
                sQty=sInvQty 
                sRate = sInvRate 
            end if
		end if
		
            %>

                <tr>
                <td class="ExcelSerial" align="center"><%=iInvsno%></td>
                <td class="ExcelDisplayCell"><%=sInvDescription%></td>
                <td class="ExcelDisplayCell" align="Right" id="tOldQty<%=iInvsno%>"><%=FormatNumber(sInvQty,2,,,0)%></td>
                <td class="ExcelDisplayCell" align="Right" id="tOldRate<%=iInvsno%>"><%=FormatNumber(sInvRate,2,,,0)%></td>
                <td class="ExcelDisplayCell" align="Right" id="tOldDis<%=iInvsno%>"><%=FormatNumber(sInvDiscount,2,,,0)%></td>
                <td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtOldAmount<%=iInvsno%>"  value="<%=FormatNumber(sInvAmount,2,,,0)%>" class="FormelemRead" size="13"></td>

                <td class="ExcelDisplayCell" align="Right" id="tQty<%=iInvsno%>"><input type="text" style="text-align: Right" NAME="txtqty<%=iInvsno%>" onkeypress="DoKeyPress('Y',12,2)" onBlur="setQty(this,'<%=iInvSno%>','Q')" value="<%=FormatNumber(sQty,3,,,0)%>" class="FormelemRead" size="13"></td>
                <td class="ExcelDisplayCell" align="Right" id="tRate<%=iInvsno%>"><input type="text" style="text-align: Right" NAME="txtRate<%=iInvsno%>" onkeypress="DoKeyPress('Y',12,2)" onBlur="setQty(txtqty<%=iInvsno%>,'<%=iInvSno%>','R')" value="<%=FormatNumber(sRate,2,,,0)%>" class="FormelemRead" size="13"></td>
                <input type="hidden" name="hRatePer" value="<%=sRatePer%>">
                </td>
                <td class="ExcelDisplayCell" align="Right" id="tDis<%=iInvsno%>">
                <input type="hidden" name="hDisPer<%=iInvSNo%>" value="<%=sDiscPer%>">
                <input type="text" style="text-align: Right" NAME="txtDis<%=iInvsno%>" value="<%=FormatNumber(sDiscount,2,,,0)%>" class="FormelemRead" size="13" onkeypress="DoKeyPress('Y',12,2)" onBlur="setQty(txtqty<%=iInvsno%>,'<%=iInvSno%>','R')">
                </td>


                <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount<%=iInvSno%>" onBlur="setTotal(this,'<%=iInvSno%>')" value="<%=FormatNumber(sAmount,2,,,0)%>" class="Formelem" size="13" onkeypress="DoKeyPress('Y',12,2)"></td>
	            <!--td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount"  value="<%=FormatNumber(sAmount,2,,,0)%>" class="FormelemRead" size="13"></td-->
                    </tr>
            <%
    Next
%>

        <tr>
        <Input type="hidden" name="hRowVal" value="<%=iInvsno%>">
    <td align="center" ></td>

    <td class="ExcelSerial" align="center"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td align="center" ></td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dInvTotal,2,,,0)%></b></td>
    <td align="right" ></td>
    <td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" readonly NAME="txtTotalInv" value="<%=FormatNumber(dinvTotal,2,,,0)%>" class="FormelemRead" size="13"></td>
     <td align="right" colspan="3"></td>
     <td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" readonly NAME="txtTotal" value="<%=FormatNumber(dTotal,2,,,0)%>" class="Formelem" size="13"></td>

        </tr>



<%
	Dim sCheckExp,CheckNode
'	dim dInvAmount
'	dInvAmount=dTotal
	Dim iCheck
	Dim sInvCatCode,sInvTaxCode,sInvTaxMode,sInvFormula,dInvTaxValue,iInvRoundOff,dInvTax,sInvTaxName
	
	For Each oInvNodEntry in oInvNodTaxRoot.childNodes
	    sInvCatCode = oInvNodEntry.getAttribute("CatCode")
	    sInvTaxCode = oInvNodEntry.getAttribute("TaxCode")
	    sInvTaxMode = oInvNodEntry.getAttribute("TaxMode")
	    sInvFormula = oInvNodEntry.getAttribute("TaxFormula")
	    dInvTaxValue = oInvNodEntry.getAttribute("TaxValue")
	    
	    For Each oNodEntry in oNodTaxRoot.childNodes
		    sCatCode=oNodEntry.Attributes.GetNamedItem("CatCode").value
		    sTaxCode=oNodEntry.Attributes.GetNamedItem("TaxCode").value
		    sTaxMode=oNodEntry.Attributes.GetNamedItem("TaxMode").value
		    sFormula=oNodEntry.Attributes.GetNamedItem("TaxFormula").value
		    dTaxValue=oNodEntry.Attributes.GetNamedItem("TaxValue").value
		    
		    
		    If sInvCatCode = "0" and sInvTaxCode = "0" and sInvTaxMode = "0" Then
			    iInvRoundOff = 0
		    Else
			    sCheckExp = "//TaxDetails/Tax[@CatCode="&sInvCatCode&" and @TaxCode="&sInvTaxCode&" and @RoundOff]"
			    Set CheckNode = oNodRoot.selectNodes(sCheckExp)
			    IF CheckNode.length <> 0 Then
				    iInvRoundOff = oInvNodEntry.Attributes.GetNamedItem("RoundOff").value
			    Else
				    iInvRoundOff  = 0
			    End IF
		    End If
		    sInvTaxName=oNodEntry.Text
		    If iInvRoundOff = 1 Then
			    dInvTax  = FormatNumber(Round(oInvNodEntry.Attributes.GetNamedItem("TaxAmount").value,0),2,,,0)
		    Else
			    dInvTax = FormatNumber(oInvNodEntry.Attributes.GetNamedItem("TaxAmount").value,2,,,0)
		    End If


		    If sCatCode = "0" and sTaxCode = "0" and sTaxMode = "0" Then
			    iRndOff = 0
		    Else
			    sCheckExp = "//TaxDetails/Tax[@CatCode="&sCatCode&" and @TaxCode="&sTaxCode&" and @RoundOff]"
			    Set CheckNode = oNodRoot.selectNodes(sCheckExp)
			    IF CheckNode.length <> 0 Then
				    iRndOff = oNodEntry.Attributes.GetNamedItem("RoundOff").value
			    Else
				    iRndOff = 0
			    End IF
		    End If
		    sTaxName=oNodEntry.Text
		    If iRndOff = 1 Then
			    dTax = FormatNumber(Round(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,0),2,,,0)
		    Else
			    dTax = FormatNumber(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,2,,,0)
		    End If
		    if Trim(sInvTaxCode)=Trim(sTaxCode) and Trim(sInvCatCode)=Trim(sCatCode) then
%>
			<tr>
				<td align="center" colspan="1"></td>
				<td class="ExcelSerial" align="right" colspan="3"><%=sInvTaxName%>&nbsp;</td>
				<%if sInvTaxMode="P" then %>
				<td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxValue%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="6" class="FormelemRead" readonly>&nbsp;%</td>
				<%else%>
				<td class="ExcelDisplayCell" align="right">
				<%
					if sInvTaxMode="K" then Response.Write "Per Pack"
					if sInvTaxMode="Q" then Response.Write "Per Qty"
				%>
				</td>
				<%end if%>
				<td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" NAME="txtOldTaxValue<%=sinvCatCode%><%=sinvTaxCode%>" value="<%=dinvTax%>"  size="11" class="FormelemRead"></td>
				<td align="center" colspan="3"></td>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  size="11" class="Formelem" onBlur="ReTotalCr()"></td>

				    </tr>
			<%
			end if 'if Trim(sInvTaxCode)=Trim(sTaxCode) and Trim(sInvCatCode)=Trim(sCatCode) then
	    next
	Next

oDOM.save server.MapPath("../temp/transaction/Voucher EntryAmd_DN_"&Session.SessionID&".xml")
%>



        <tr>
        <td align="center" colspan="1"></td>
    <td class="ExcelSerial" align="right" colspan="4"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"> <input type="text" style="text-align: Right" NAME="txtFInvValue"  size="13" value="<%=FormatNumber(dInvInvAmount,2,,,0)%>" class="FormelemRead"></td>
    <td align="right" colspan="3"></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="Formelem">

    </td>
        </tr>

        <tr>
        <td align="center" colspan="1"></td>
        <td class="ExcelSerial" align="right" colspan="8"><b>Credit Note Value&nbsp; </b></td>
    <!--td class="ExcelSerial" align="right" colspan="3"><b>Credit Note Value&nbsp; </b></td-->
    <td class="ExcelDisplayCell" align="right" id="tDrVal"> <input type="text" style="text-align: Right" NAME="txtCrNoteValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="FormelemRead" readonly>

    </td>
        </tr>

        <tr>
        </tr>


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

			<Textarea name="txtNarration" class="FormElem" cols="40" rows="4"><%=sNarration%></Textarea>

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
                                                                <input type="button" value="Cancel" name="B6" class="ActionButton" onClick="Cancel('CreditVouchers.ASP')" >
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
