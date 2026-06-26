<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNPurInvEntry.asp.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	October 20, 2004
	'Modified By				:	Ragavendran r
	'Modified On				:	Feb 06,2010
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
Dim sFromApp
Dim sFinPeriod,sFinFrm,sFinTo,sValTemp,sCallFrom,iPurInvoiceNo

sFinPeriod = Session("FinPeriod")
sValTemp = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp(0))
sFinTo = Trim(sValTemp(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

sInvoiceNo=Request.Form("selInvoiceNo")
sBookName=Request.Form("hBookName")
sFromApp = Request.Form("hFromApp")
sCallFrom = Request.QueryString("hCallFrom")
sTemp = Split(sInvoiceNo,":")
sInvoiceNo = sTemp(0)

iInvNo = sInvoiceNo

sQuery = "Select FromApplication,OtherApplnTransNo From Acc_T_CreatedVoucherHeader "&_
		 "Where CreatedTransNo = "&iInvNo&" and FromApplication is Not NULL "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sFromPur = "Y"
	iPurInvoiceNo = objRs(1)
Else
	sFromPur = "N"
	iPurInvoiceNo = 0
End IF
objRs.Close


sUserid = getUserID()

Dim sRetVal
'oDOM.load server.MapPath("../xmldata/Voucher/"&sInvoiceNo&".xml")
sRetVal = GetVouchXML(sInvoiceNo)
oDOM.Load server.MapPath(sRetVal)


set oNodRoot=oDOM.documentElement

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
			if oNodEntry.nodeName="PurInvoice" then
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

sInvValue=oNodTaxRoot.Attributes.Item(2).nodeValue
dInvAmount=oNodTaxRoot.Attributes.Item(0).nodeValue

dim dInvAmount
'dInvAmount = sInvValue

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML id="TaxData" src="<%="../temp/transaction/Voucher Entry_DN_"&Session.SessionID&".xml"%>"></XML>
<XML id="GJVoucher"></XML>
<SCRIPT LANGUAGE=javascript SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<script language="vbscript" >
Function SaveXML()
	Dim sExp,TempNode,sCheckVal,iCtr
	dim iCrAccHead,sDesc,oDOMGJ,oDGjRoot,oDGjEntry,oDGjAcc,oDGjNarr,oDSubNode
	Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newE

	if trim(document.formname.hCallFrom.value)="CR" then

			set RootNode=TaxData.documentElement
			For Each oNodTemp in RootNode.childNodes
				if oNodTemp.nodeName="Details" then
				 	oNodTemp.Attributes.GetNamedItem("VouDate").value=document.formname.ctlDate.GetDate
				end if
			next

			sExp = "//PurInvoice"
			Set TempNode = RootNode.SelectNodes(sExp)

			Set newElem = TaxData.CreateElement("Narration")
			newElem.Text = document.formname.txtNarration.value
			RootNode.appendChild newElem
			
			Set newElem = TaxData.CreateElement("PurchaseInvoiceEntry")
	        newElem.setAttribute "InvoiceNo",document.formname.hInvoiceNo.value
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

				Set newElem  = TaxData.createAttribute("PurTransNo")
				newElem.value = document.formname.hdTransNo.value
				TempNode.Item(0).setAttributeNode(newElem)

				Set newElem  = TaxData.createAttribute("CRNoteType")
				newElem.value = document.formname.SelCrAgain.value
				TempNode.Item(0).setAttributeNode(newElem)


			End IF

			IF document.formname.SelCrAgain.value = "A" Then
				sExp = "//TaxDetails"
				Set TempNode = RootNode.SelectNodes(sExp)
				IF TempNode.length <> 0 Then
					TempNode.Item(0).Attributes.Item(0).nodeValue = document.formname.txtCrNoteValue.value
				End IF
			End IF

			IF document.formname.SelCrAgain.value = "A" Then
				sExp = "//Entry"
				Dim dEachItmVal
				Set TempNode = RootNode.SelectNodes(sExp)
				IF TempNode.length <> 0 Then
					For iCtr = 0 To TempNode.length - 1
						Set dEachItmVal = Eval("document.formname.txtAmount"&iCtr+1)
						TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value = dEachItmVal.Value
					Next
				End IF

				sExp = "//Tax"
				Set TempNode = RootNode.SelectNodes(sExp)
				IF TempNode.length <> 0 Then
					For iCtr = 0 To TempNode.length - 1
						TempNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value = "0.00"
					Next
				End IF

			End IF

			'MsgBox "1"
			IF document.formname.SelAccountHd.value = "G" Then
				iCrAccHead = document.formname.hCrAccHead.value
				sDesc = document.all.spAccHead.innerHTML
				sExp = "//AccHead"
				Set TempNode = RootNode.selectNodes(sExp)
				'MsgBox TempNode.length
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

			'alert TaxData.xml

			IF CheckFinDate Then
				set objhttp = CreateObject("Microsoft.XMLHTTP")
				objhttp.Open "POST","XMLSave.asp?Mod=DN&Name=Voucher Entry", false
				objhttp.send TaxData.XMLDocument
				if objhttp.responseText <> "" then
					Msgbox(objhttp.responseText)
				else
					'alert(TaxData.xml)
					'MsgBox "OK "
					document.formname.B2.disabled = True
					document.formname.action="VouCNPurInvUpdate.asp"
					document.formname.submit()
				end if
			End IF

	else ' if trim(document.formname.hCallFrom.value)="CN" then

			set oNodRoot = TaxData.documentElement
			set oDGjRoot = GJVoucher.createElement("voucher")
			GJVoucher.appendChild oDGjRoot
			for each oNodHeader in oNodRoot.childNodes
				if oNodHeader.nodeName = "Header" then
					for each oDSubNode in oNodHeader.childNodes
						if oDSubNode.nodeName="Organization" then
							oDGjRoot.setAttribute "UnitNo",oDSubNode.getAttribute("OrgId")
							oDGjRoot.setAttribute "UnitName",oDSubNode.text
							oDGjRoot.setAttribute "BookNo",document.formname.hBookCode.value
							oDGjRoot.setAttribute "BookName",document.formname.hBookName.value
							oDGjRoot.setAttribute "CRDR","D"
						elseif oDSubNode.nodeName="Party" then
							oDGjRoot.setAttribute "PartyCode",oDSubNode.getAttribute("ParType")&"?"&oDSubNode.getAttribute("ParSubType")&"?"& oDSubNode.getAttribute("ParSubTypeName")&"?"&oDSubNode.getAttribute("ParCode")
							oDGjRoot.setAttribute "Approver",document.formname.selUserId.value
							oDGjRoot.setAttribute "PartyName",oDSubNode.text
						elseif oDSubNode.nodeName="PurInvoice" then
							oDGjRoot.setAttribute "InvNo",oDSubNode.getAttribute("PurInvNo")
							oDGjRoot.setAttribute "InvDate",oDSubNode.getAttribute("PurInvDate")
						end if
					next
				elseif oNodHeader.nodeName="Details" then
				oDGjRoot.setAttribute "VouDate",document.formname.ctlDate.GetDate
				for each oNodEntry in oNodHeader.childNodes
						if oNodEntry.nodeName="Entry" then
							set oDGjEntry= GJVoucher.createElement("Entry")
							iCnt = iCnt + 1
							oDGjEntry.setAttribute "No",iCnt
							oDGjEntry.setAttribute "CRDR","C"
							oDGjEntry.setAttribute "Payto",oNodEntry.getAttribute("PayTo")
							oDGjEntry.setAttribute "Amount",document.formname.txtCrNoteValue.value
							oDGjEntry.setAttribute "AccUnit",""
							oDGjEntry.setAttribute "AccName",""
							oDGjEntry.setAttribute "TDSAmount",""
							oDGjEntry.setAttribute "TDAElgi","0"
							oDGjEntry.setAttribute "TDSPercentage","0"
							oDGjEntry.setAttribute "PayRecAmount","0"
							oDGjRoot.appendChild oDGjEntry
							for each oNodDeatils in oNodEntry.childNodes
								if oNodDeatils.nodeName="AccHead" then
									set oDGjAcc = GJVoucher.createElement("AccHead")
									oDGjAcc.setAttribute "No",oNodDeatils.getAttribute("No")
									oDGjAcc.setAttribute "CostCenter",oNodDeatils.getAttribute("CostCenter")
									oDGjAcc.setAttribute "Analytical",oNodDeatils.getAttribute("Analytical")
									oDGjAcc.setAttribute "Name",oNodDeatils.getAttribute("Name")
									oDGjAcc.setAttribute "Type",oNodDeatils.getAttribute("Type")
									oDGjAcc.setAttribute "TransFlag","W"
									oDGjEntry.appendChild oDGjAcc
								end if
								Set oDGjNarr = GJVoucher.CreateElement("Narration")
								oDGjNarr.Text = ""
								oDGjEntry.appendChild oDGjNarr
							next
						end if
						exit for ' for only one Record should be inserted
					next
				end if
			next



			IF CheckFinDate Then
				set objhttp = CreateObject("Microsoft.XMLHTTP")
				objhttp.Open "POST","XMLSave.asp?Mod=CNGJ&Name=Voucher Entry", false
				objhttp.send GJVoucher.XMLDocument
				if objhttp.responseText <> "" then
					Msgbox(objhttp.responseText)
				else
					'alert(TaxData.xml)
					'MsgBox "OK "
					document.formname.B2.disabled = True
					document.formname.action="VouCNGJGenerate.asp?hCallFrom=PI"
					document.formname.submit()
				end if
			End IF

	end if 'if trim(document.formname.hCallFrom.value)="CN" then

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
	Dim sVal,sPreInvVal,sNewInvVal,dDiffVal,dRowLen
	sVal = sObj.value
	dRowLen = document.formname.hRowVal.value

	IF sCallType = "1" Then
		document.formname.reset()
		For iCtr = 0 To sObj.length - 1
			IF sObj.Options(iCtr).value = sVal Then
				sObj.selectedIndex = iCtr
				Exit For
			End IF
		Next
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
			Eval("document.all.tDis"&iCtr).class = "ExcelDisplayCell"

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
	document.formname.txtCrNoteValue.value = dCrVal

	sExp = "//TaxDetails"
	Set TempNode = RootNode.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.item(0).Value = dCrVal
	End IF

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
'==========================================================================================================================

</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" >
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
<input type="hidden" name="hCrAccHead" value="0">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hInvoiceNo" value="<%=iPurInvoiceNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%IF CStr(sFromApp) <> "C" Then %>
				Debit Note Other Invoices
		<%Else%>
				Credit Note Purchase Invoices
		<%End IF %>
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
                                                    <Select name="SelCrAgain" class="FormElem" onChange="SetRetVal(this,'1')">
                                                    <Option Value="0">Select</Option>
                                                    <Option Value="Q">Quantity</Option>
                                                    <Option Value="R">Rate</Option>
                                                    <Option Value="D">Discount</Option>
                                                    <Option Value="A">Quality</Option>

                                                    </Select>
                                                    </td>

                                                        </tr>
                                                        <tr>
														<td class="FieldCellSub" width="200">Select Account Head</td>
														<td class="FieldCell" width="200">
														<Select name="SelAccountHd" class="FormElem" onChange="AccHead(this)">
															<Option Value="0">ITEM ACCOUNT HEAD</Option>
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
    <td class="ExcelHeaderCell" align="center" width="55">Value</td>
    <td class="ExcelHeaderCell" align="center" width="55">Discount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Amount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="55">Value</td>
    <td class="ExcelHeaderCell" align="center" width="55">Discount</td>
    <td class="ExcelHeaderCell" align="center" width="55">Amount</td>
        </tr>
<%
	Dim sRatePer
	For Each oNodEntry in oNodDeatils.childNodes
		iSno=oNodEntry.Attributes.GetNamedItem("No").value
		sDescription=oNodEntry.Attributes.GetNamedItem("PayTo").value
		sAmount=oNodEntry.Attributes.GetNamedItem("Amount").value
		sQty=oNodEntry.Attributes.GetNamedItem("Qty").value
		sValue=oNodEntry.Attributes.GetNamedItem("ActValue").value
		'oNodEntry.Attributes.GetNamedItem("Rate").value=CDbl(oNodEntry.Attributes.GetNamedItem("Amount").value)/CDbl(sQty)
		sRate = oNodEntry.Attributes.GetNamedItem("Rate").value
		'sRatePer = oNodEntry.Attributes.GetNamedItem("RatePer").value
		sRatePer = 1
		sDiscPer =oNodEntry.Attributes.GetNamedItem("DisPer").value
		sDiscount=oNodEntry.Attributes.GetNamedItem("DisAmount").value

		dTotal=CDbl(dTotal)+CDbl(sAmount)

%>

    <tr>
    <td class="ExcelSerial" align="center"><%=isno%></td>
    <td class="ExcelDisplayCell"><%=sDescription%></td>
    <td class="ExcelDisplayCell" align="Right" id="tOldQty<%=isno%>"><%=FormatNumber(sQty,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" id="tOldRate<%=isno%>"><%=FormatNumber(sRate,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" id="tOldDis<%=isno%>"><%=FormatNumber(sDiscount,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount"  value="<%=FormatNumber(sAmount,2,,,0)%>" class="FormelemRead" size="13"></td>

    <td class="ExcelDisplayCell" align="Right" id="tQty<%=isno%>"><input type="text" style="text-align: Right" NAME="txtqty<%=isno%>" onBlur="setQty(this,'<%=iSno%>','Q')" value="<%=FormatNumber(sQty,3,,,0)%>" class="FormelemRead" size="13"></td>
    <td class="ExcelDisplayCell" align="Right" id="tRate<%=isno%>"><input type="text" style="text-align: Right" NAME="txtRate<%=isno%>" onBlur="setQty(txtqty<%=isno%>,'<%=iSno%>','R')" value="<%=FormatNumber(sRate,2,,,0)%>" class="FormelemRead" size="13"></td>
    <input type="hidden" name="hRatePer" value="<%=sRatePer%>">
    </td>
    <td class="ExcelDisplayCell" align="Right" id="tDis<%=isno%>">
    <input type="hidden" name="hDisPer<%=iSNo%>" value="<%=sDiscPer%>">
    <input type="text" style="text-align: Right" NAME="txtDis<%=isno%>" value="<%=FormatNumber(sDiscount,2,,,0)%>" class="FormelemRead" size="13" onBlur="setQty(txtqty<%=isno%>,'<%=iSno%>','R')">
    </td>


    <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount<%=iSno%>" onBlur="setTotal(this,'<%=iSno%>')" value="<%=FormatNumber(sAmount,2,,,0)%>" class="Formelem" size="13"></td>
	<!--td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount"  value="<%=FormatNumber(sAmount,2,,,0)%>" class="FormelemRead" size="13"></td-->
        </tr>
<%
	next
%>

        <tr>
        <Input type="hidden" name="hRowVal" value="<%=isno%>">
    <td align="center" ></td>

    <td class="ExcelSerial" align="center"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td align="center" ></td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
    <td align="right" ></td>
    <td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" readonly NAME="txtTotalInv" value="<%=FormatNumber(dTotal,2,,,0)%>" class="FormelemRead" size="13"></td>
     <td align="right" colspan="3"></td>
     <td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" readonly NAME="txtTotal" value="<%=FormatNumber(dTotal,2,,,0)%>" class="Formelem" size="13"></td>

        </tr>



<%
	Dim sCheckExp,CheckNode
'	dim dInvAmount
'	dInvAmount=dTotal
	Dim iCheck
	For Each oNodEntry in oNodTaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.GetNamedItem("CatCode").value
		sTaxCode=oNodEntry.Attributes.GetNamedItem("TaxCode").value
		sTaxMode=oNodEntry.Attributes.GetNamedItem("TaxMode").value
		sFormula=oNodEntry.Attributes.GetNamedItem("TaxFormula").value
		dTaxValue=oNodEntry.Attributes.GetNamedItem("TaxValue").value


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
		if trim(oNodEntry.Attributes.GetNamedItem("TaxAmount").value)<>"" then
			If iRndOff = 1 Then
				dTax = FormatNumber(Round(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,0),2,,,0)
			Else
				dTax = FormatNumber(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,2,,,0)
			End If
		end if
%>
			<tr>
				<td align="center" colspan="1"></td>
				<td class="ExcelSerial" align="right" colspan="3"><%=sTaxName%>&nbsp;</td>
				<%if sTaxMode="P" then %>
				<td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxValue%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="6" class="FormelemRead" readonly>&nbsp;%</td>
				<%else%>
				<td class="ExcelDisplayCell" align="right">
				<%
					if sTaxMode="K" then Response.Write "Per Pack"
					if sTaxMode="Q" then Response.Write "Per Qty"
				%>
				</td>
				<%end if%>
				<td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue" value="<%=dTax%>"  size="11" class="FormelemRead"></td>
				<td align="center" colspan="3"></td>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  size="11" class="Formelem" onBlur="ReTotalCr()"></td>

				    </tr>

			<%
	next


oDOM.save server.MapPath("../temp/transaction/Voucher Entry_DN_"&Session.SessionID&".xml")

%>



        <tr>
        <td align="center" colspan="1"></td>
    <td class="ExcelSerial" align="right" colspan="4"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"> <input type="text" style="text-align: Right" NAME="txtFInvValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="FormelemRead"></td>
    <td align="right" colspan="3"></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="Formelem">

    </td>
        </tr>

        <tr>
        <td align="center" colspan="1"></td>
        <td class="ExcelSerial" align="right" colspan="8"><b>Credit Note Value&nbsp; </b></td>
    <!--td class="ExcelSerial" align="right" colspan="3"><b>Credit Note Value&nbsp; </b></td-->
    <td class="ExcelDisplayCell" align="right" id="tDrVal"> <input type="text" style="text-align: Right" NAME="txtCrNoteValue"  size="13" value="0.00" class="FormelemRead" readonly>

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

			<Textarea name="txtNarration" class="FormElem" cols="40" rows="4"></Textarea>

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
