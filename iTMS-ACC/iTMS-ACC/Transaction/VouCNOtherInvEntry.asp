<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNOtherInvEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	October 20, 2004
	'Modified By				:	Ragavendran R
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
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff,sFromSal
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId,sTemp,iitemCode,iClassCode,iSalRetQty
Dim sFinFromDate,sFinTodate

Dim sFinPeriod,sFinFrm,sFinTo,sValTemp
Dim sCallFrom ,iCnt,sBookNumber,sInvFrom
Dim oDOMGJ,oDGjRoot,oDGjEntry,oDGjAcc,oDGjNarr,oDSubNode,nRatePer,iSalInvoiceNo
sFinPeriod = Session("FinPeriod")
sValTemp = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp(0))
sFinTo = Trim(sValTemp(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"
sFinFromDate = "01/04/"&Trim(sValTemp(0))
sFinTodate = "31/03/"&Trim(sValTemp(1))
iCnt =0
sBookNumber = Request.Form("selBook")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set oDOMGJ = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

sInvoiceNo=Request.Form("selInvoiceNo")
sTemp = Split(sInvoiceNo,":")
sBookName=Request.Form("hBookName")
'Response.Write sInvoiceNo
sCallFrom= Request.QueryString("hCallFrom")
sInvoiceNo = sTemp(0)
iInvNo = sInvoiceNo

sUserid = getUserID()

sQuery = "Select FromApplication,OtherApplnTransNo From Acc_T_CreatedVoucherHeader "&_
		 "Where CreatedTransNo = "&iInvNo&" and FromApplication is Not NULL "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sFromSal = "Y"
	iSalInvoiceNo = objRs(1)
Else
	sFromSal = "N"
	iSalInvoiceNo = 0
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

sInvValue=oNodTaxRoot.Attributes.Item(2).nodeValue
dInvAmount=oNodTaxRoot.Attributes.Item(2).nodeValue

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
<xml id="GLData"><Root></Root></xml>
<SCRIPT LANGUAGE=javascript SRC="../scripts/VouSalesReturnOthInv.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<script language="vbscript" >
'===========================================================
Function TaxDetail()
	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
	
End Function
'----------------------------------------------
Function Init()
Dim stemp,sTempValue,sTempObj
Dim sVal,sPreInvVal,sNewInvVal,dDiffVal,sLen,sTotQty
	if document.formname.hCallFromVoucher.value="GJ" then
		stemp = InoviceNo.innerText 
		sTempValue =  split(stemp,"-")
		document.formname.txtNarration.value = "CR Note for "& sTempValue(0) & " "& sTempValue(1)
	end if
	
	if document.formname.hInvCallFrom.value="SR" then
		
		sVal = document.formname.SelCrAgain.value 
		sLen = document.formname.hRowVal.value
		
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
		
		IF CStr(document.formname.hAccType.Value) = "C" Then
			dDiffVal = CDbl(dPreInvVal) - CDbl(sNewInvVal)
		Else
			dDiffVal = CDbl(sNewInvVal) - CDbl(dPreInvVal) 
		End IF
		
		dDiffVal = FormatNumber(dDiffVal,2,,,0)
		sNewInvVal = FormatNumber(sNewInvVal,2,,,0)
		
		For iCtr = 1 To sLen
			IF CStr(sVal) = "Q" Then
				Eval("document.formname.txtQty"&iCtr).className = "FormElem"
				Eval("document.all.tQty"&iCtr).className = "ExcelInputCell"
				Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
				Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
				Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
				Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"
				
				document.formname.txtCrNoteValue.value = sNewInvVal
				document.formname.txtCrNoteValue.readOnly = True
				document.formname.txtCrNoteValue.className = "FormElemRead"
				document.all.tCrValue.className = "ExcelDisplayCell"
				
				Eval("document.formname.txtdis"&iCtr).readonly = True
				Eval("document.formname.txtQty"&iCtr).readonly = False
				Eval("document.formname.txtRate"&iCtr).readonly = True
				
				sTotQty = sTotQty &  Eval("document.formname.txtQty"&iCtr).value
			End IF
		Next
	
		document.formname.txtNarration.value = "CR Note for "& split(InoviceNo.innerText,"-")(0) &" "& split(InoviceNo.innerText,"-")(1) & " Sales Return Qty: "&	sTotQty
			
	end if ' if document.formname.hInvCallFrom.value="SR" then
	
	sFromDate = document.formname.hFromDate.value
	sTodate = document.formname.hToDate.value
	if DateDiff("d",sTodate,date)>0 then
	    document.formname.ctlDate.setMinDate = sFromDate
	    document.formname.ctlDate.setMaxDate = sToDate
	    document.formname.ctlDate.setDate = sToDate
	else
	    document.formname.ctlDate.setMinDate = sFromDate
	    document.formname.ctlDate.setMaxDate = date
	    document.formname.ctlDate.setDate =date
	end if 

End Function
'---------------------------------------------
Function SaveXML()

	Dim sExp,TempNode,sCheckVal,iCrAccHead,sDesc,sTempUnitNo,sTempUnitName
	Dim oDOMGJ,oDGjRoot,oDGjEntry,oDGjAcc,oDGjNarr,oDSubNode,sPartyName
	Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1
	Dim NodPayRec,NodDoc,NodRecCount,sPartyNo,sInvNoDet,sInvDate,sTempArr,sTempValue
	Dim sPartyDet,sArrValue,sRetVal2,setFlag,sTempAccHead,TaxRoot
	set RootNode=TaxData.documentElement
	setFlag = false
	'alert(document.formname.hCallFromVoucher.value)
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
	
	Set newElem = TaxData.CreateElement("SalesInvoiceEntry")
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
		
		Set newElem  = TaxData.createAttribute("SalTrNo")
		newElem.value = document.formname.hdTransNo.Value
		TempNode.Item(0).setAttributeNode(newElem)
		
	End IF
	
	sExp = "//TaxDetails"
	Set TempNode = RootNode.SelectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.Item(0).nodeValue = document.formname.txtCrNoteValue.Value
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
					sPartyDet = sTempUnitNo &"&ParSubType="&oDSubNode.getAttribute("ParSubType")&"&ParType=" &oDSubNode.getAttribute("ParType")&"&PartyCode="&oDSubNode.getAttribute("ParCode")
					sPartyName = oDSubNode.text
				elseif oDSubNode.nodeName="SaleInvoice" then
					sInvNoDet ="SALE INV NO:"&oDSubNode.getAttribute("InvNo") &" Dt:"& oDSubNode.getAttribute("InvDate") 
					sInvDate = oDSubNode.getAttribute("InvDate") 
				end if
			next
		elseif strcomp(oNodHeader.nodeName,"Details")=0 then
			oDGjRoot.setAttribute "VouDate",document.formname.ctlDate.GetDate
			oDGjRoot.setAttribute "BookAcchead",sBookACHead
			oDGjRoot.setAttribute "Approver",document.formname.selUserId.value
			iItemCount = 0
			for each oNodEntry in oNodHeader.childNodes
				if oNodEntry.nodeName="Entry" then
				
				
					if setFlag = False then
						'First Entry
						set oDGjEntry= GJVoucher.createElement("Entry")
						iCnt = iCnt + 1
						oDGjEntry.setAttribute "No",iCnt
						oDGjEntry.setAttribute "CRDR","C"
						oDGjEntry.setAttribute "Payto","0"
						oDGjEntry.setAttribute "Amount",document.formname.txtCrNoteValue.Value
						oDGjEntry.setAttribute "AccUnit",sTempUnitNo 
						oDGjEntry.setAttribute "AccName",sTempUnitName 
						oDGjEntry.setAttribute "TdsAmount","0.00"
						oDGjEntry.setAttribute "TDSElgi","0"
						oDGjEntry.setAttribute "TdsPercentage","0"
						oDGjEntry.setAttribute "PayRecAmount","0"
						oDGjRoot.appendChild oDGjEntry 
						for each oNodDeatils in oNodEntry.childNodes
							if oNodDeatils.nodeName="AccHead" then
								set objhttp = CreateObject("Microsoft.XMLHTTP")
								objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sPartyDet, false
								objhttp.send

								IF objhttp.responseText <> "" Then
									sRetVal2 = objhttp.responseText
									sArrValue = split(sRetVal2,":")
								End IF
								
								set oDGjAcc = GJVoucher.createElement("AccHead")
								oDGjAcc.setAttribute "No",sPartyNo 
								oDGjAcc.setAttribute "Pay",sArrValue(0)
								oDGjAcc.setAttribute "Rec",sArrValue(1)
								oDGjAcc.setAttribute "Name",sPartyName 
								oDGjAcc.setAttribute "Type","P"
								oDGjAcc.setAttribute "Adv",sArrValue(2)
								oDGjEntry.appendChild oDGjAcc 
							end if
							
							set objhttp = CreateObject("Microsoft.XMLHTTP")
							objhttp.Open "GET","GetGJXML.asp?hTransNo="&document.formname.hdTransNo.value, false
							objhttp.send 

								if objhttp.responseText <> "" then
									sTempArr = objhttp.responseText
									sTempValue= split(sTempArr,"#")
									if cdbl(sTempValue(3)) > cdbl(sTempValue(5)) then
										set NodPayRec = GJVoucher.createElement("PayRec")
										oDGjEntry.appendChild NodPayRec 
									
										set NodDoc = GJVoucher.createElement("Doc")
											NodDoc.setAttribute "No",sTempValue(0)
											NodDoc.setAttribute "InvNo",sTempValue(1)
											NodDoc.setAttribute "InvDate",sTempValue(2)
											NodDoc.setAttribute	"TransAmount",sTempValue(3)
											NodDoc.setAttribute "AmtAdjusted",sTempValue(4)
											NodDoc.setAttribute "AmtToAdjust", document.formname.txtCrNoteValue.Value 
											NodDoc.setAttribute "DocType",sTempValue(6)
											NodDoc.setAttribute "AmtToAccount",sTempValue(5)
											NodDoc.setAttribute "PayableNo",sTempValue(8)
											NodDoc.setAttribute "AdjType",sTempValue(9)
										NodPayRec.appendChild NodDoc 
									end if 'if cdbl(sTempValue(3))> cdbl(sTempValue(4)) then
								end if'if objhttp.responseText <> "" then
								
							set NodRecCount = GJVoucher.createElement("RecCount")
							NodRecCount.setAttribute "Val","1"
							oDGjEntry.appendChild NodRecCount 
							Set oDGjNarr = GJVoucher.CreateElement("Narration")
							oDGjNarr.Text = document.formname.txtNarration.value
							oDGjEntry.appendChild oDGjNarr
						next
						setFlag = True
					end if ' 	if setFlag = False then
						
					'Second Entry
					set oDGjEntry= GJVoucher.createElement("Entry")
					iCnt = iCnt + 1
					iItemCount = iItemCount + 1
					oDGjEntry.setAttribute "No",iCnt
					oDGjEntry.setAttribute "CRDR","D"
					oDGjEntry.setAttribute "Payto",""
					oDGjEntry.setAttribute "Amount",document.formname.txtTotal.Value
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
							sTempAccHead = oNodDeatils.getAttribute("No")
							oDGjAcc.setAttribute "No",oNodDeatils.getAttribute("No")
							oDGjAcc.setAttribute "CostCenter",oNodDeatils.getAttribute("CostCenter")
							oDGjAcc.setAttribute "Analytical",oNodDeatils.getAttribute("Analytical")
							oDGjAcc.setAttribute "Name",oNodDeatils.getAttribute("Name")
							oDGjAcc.setAttribute "Type",oNodDeatils.getAttribute("Type")
							if oNodDeatils.getAttribute("No")<>"0" then
								oDGjAcc.setAttribute "TransFlag","A"	
							else
								oDGjAcc.setAttribute "TransFlag","W"
							end if	
							oDGjEntry.appendChild oDGjAcc 
						end if
						Set oDGjNarr = GJVoucher.CreateElement("Narration")
						oDGjNarr.Text = oNodEntry.getAttribute("PayTo") & "&"& eval("document.formname.txtQty"&iItemCount).value &"&"& oNodEntry.getAttribute("UOMValue") &"&"& cdbl(eval("document.formname.txtRate"&iItemCount).value)/oNodEntry.getAttribute("RatePer")
						oDGjEntry.appendChild oDGjNarr
					next
				end if
				exit for ''added by ragav on Sep 27 for avoid the Multiple Entry
			next
		elseif strcomp(oNodHeader.nodeName,"TaxDetails")=0 then
			set TaxRoot = oNodHeader
		end if
	next
	
	
			dInvAmount = document.formname.txtTotal.value
			dBasicTotal = dInvAmount
			dTotal	= dInvAmount
'	alert(TaxRoot.xml)
					For Each oNodEntry in TaxRoot.childNodes
						sCatCode=oNodEntry.Attributes.Item(0).nodeValue 
						sTaxCode=oNodEntry.Attributes.Item(1).nodeValue 
						sTaxMode=oNodEntry.Attributes.Item(2).nodeValue 
						sFormula=oNodEntry.Attributes.Item(3).nodeValue 
						dTaxValue=oNodEntry.Attributes.Item(5).nodeValue
						sTempAccHead = oNodEntry.Attributes.Item(6).nodeValue 
					'	alert(stempAccHead)
					'	alert(dTaxValue)
					if sTempAccHead <>"0" then
						if IsObject(eval("document.formname.txtTaxValue"&sCatCode&sTaxCode)) then
						
							dTax = eval("document.formname.txtTaxValue"&sCatCode&sTaxCode).value
							if dTax <> 0 then
							
								set oDGjEntry= GJVoucher.createElement("Entry")
								iCnt = iCnt + 1
								oDGjEntry.setAttribute "No",iCnt
							
								if cdbl(dTax)< 0 then
									oDGjEntry.setAttribute "CRDR","C"
								else
									oDGjEntry.setAttribute "CRDR","D"
								end if
								dTax = round(dtax,2)
								oDGjEntry.setAttribute "Payto",""
								oDGjEntry.setAttribute "Amount",Abs(dtax)
								oDGjEntry.setAttribute "AccUnit",sTempUnitNo 
								oDGjEntry.setAttribute "AccName",sTempUnitName 
								oDGjEntry.setAttribute "TdsAmount","0.00"
								oDGjEntry.setAttribute "TDSElgi","0"
								oDGjEntry.setAttribute "TdsPercentage","0"
								oDGjEntry.setAttribute "PayRecAmount","0"
								oDGjRoot.appendChild oDGjEntry 
							
								set oDGjAcc = GJVoucher.createElement("AccHead")
								oDGjAcc.setAttribute "No",sTempAccHead 
								oDGjAcc.setAttribute "CostCenter","0"
								oDGjAcc.setAttribute "Analytical","0"
								oDGjAcc.setAttribute "Name",""
								oDGjAcc.setAttribute "Type","G"
								oDGjAcc.setAttribute "TransFlag","A"	
								oDGjEntry.appendChild oDGjAcc 
						
								Set oDGjNarr = GJVoucher.CreateElement("Narration")
								oDGjNarr.Text = oNodEntry.text
								oDGjEntry.appendChild oDGjNarr
							end if 'if cdbl(dTaxValue) <> cdbl(0) then
						end if	
					else
						for each subnode in oDGjRoot.childNodes
							if strcomp(subnode.nodeName,"Entry") = 0 then
								if strcomp(subnode.getAttribute("No"),"2")= 0 then
									dTempAmount = subnode.getAttribute("Amount")
									dTempAmount = cdbl(dtempAmount) + cdbl(dTaxValue)
									subnode.setAttribute "Amount",dtempAmount
								end if 
							end if
						next
					end if'	if sTempAccHead <>"0" then
				next ' For TaxRoot.childNodes
	
				
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
	
	IF CStr(document.formname.hAccType.Value) = "C" Then
		dDiffVal = CDbl(dPreInvVal) - CDbl(sNewInvVal)
	Else
		dDiffVal = CDbl(sNewInvVal) - CDbl(dPreInvVal) 
	End IF
	
	dDiffVal = FormatNumber(dDiffVal,2,,,0)
	sNewInvVal = FormatNumber(sNewInvVal,2,,,0)
	
	
	
	For iCtr = 1 To sLen
		IF CStr(sVal) = "Q" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElem"
			Eval("document.all.tQty"&iCtr).className = "ExcelInputCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"
			
			document.formname.txtCrNoteValue.value = sNewInvVal
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tCrValue.className = "ExcelDisplayCell"
			
			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = False
			Eval("document.formname.txtRate"&iCtr).readonly = True
			
			sTotQty = sTotQty &  Eval("document.formname.txtQty"&iCtr).value
		End IF
	
		IF CStr(sVal) = "R" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElem"
			Eval("document.all.tRate"&iCtr).className = "ExcelInputCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"
			
			document.formname.txtCrNoteValue.value = sNewInvVal
			'document.formname.txtCrNoteValue.value = dDiffVal
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tCrValue.className = "ExcelDisplayCell"
			
			Eval("document.formname.txtdis"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = False
			Eval("document.formname.txtQty"&iCtr).readonly = True
		End IF
	
		IF CStr(sVal) = "D" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElem"
			Eval("document.all.tDis"&iCtr).className = "ExcelInputCell"
			
			document.formname.txtCrNoteValue.value = sNewInvVal
			'document.formname.txtCrNoteValue.value = dDiffVal
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tCrValue.className = "ExcelDisplayCell"
			
			Eval("document.formname.txtdis"&iCtr).readonly = False
			Eval("document.formname.txtRate"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = True
			
		End IF
	
		IF CStr(sVal) = "0" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"
			
			document.formname.txtCrNoteValue.readOnly = True
			document.formname.txtCrNoteValue.className = "FormElemRead"
			document.all.tCrValue.className = "ExcelDisplayCell"
			
			Eval("document.formname.txtDis"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = True
			
		End IF
		
		IF CStr(sVal) = "A" Then
			Eval("document.formname.txtQty"&iCtr).className = "FormElemRead"
			Eval("document.all.tQty"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtRate"&iCtr).className = "FormElemRead"
			Eval("document.all.tRate"&iCtr).className = "ExcelDisplayCell"
			Eval("document.formname.txtDis"&iCtr).className = "FormElemRead"
			Eval("document.all.tDis"&iCtr).className = "ExcelDisplayCell"
		
			document.formname.txtCrNoteValue.readOnly = False
			document.formname.txtCrNoteValue.className = "FormElem"
			document.all.tCrValue.className = "ExcelInputCell"
		
			Eval("document.formname.txtDis"&iCtr).readonly = True
			Eval("document.formname.txtRate"&iCtr).readonly = True
			Eval("document.formname.txtQty"&iCtr).readonly = True
			
			ResetTax()
		End IF
	Next

	if trim(sVal)="Q" then
		document.formname.txtNarration.value = "CR Note for "& split(InoviceNo.innerText,"-")(0) &" "& split(InoviceNo.innerText,"-")(1) & " Against Quantity: "&	sTotQty
	else
		document.formname.txtNarration.value = "CR Note for "& split(InoviceNo.innerText,"-")(0) &" "& split(InoviceNo.innerText,"-")(1) 
	end if
	
	if document.formname.hInvCallFrom.value="SR" then
		document.formname.txtNarration.value = "CR Note for "& split(InoviceNo.innerText,"-")(0) &" "& split(InoviceNo.innerText,"-")(1) & " Sales Return Qty: "&	sTotQty
	end if

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

	set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),GLData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	if OutValue.getAttribute("Action")="CLOSE" then exit function
    sQuery = OutValue.getAttribute("PassQuery")
	while OutValue.getAttribute("Action")<>"Done"
		set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?"&sQuery,GLData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		if OutValue.getAttribute("Action")="CLOSE" then exit function
		sQuery = OutValue.getAttribute("PassQuery")
	wend

set ndRoot = GLData.documentElement
    if ndRoot.hasChildNodes() then
        for each ndChild in ndRoot.childNodes
            iAccCode = ndChild.getAttribute("RetField0")
	        sDesc = ndChild.getAttribute("RetField5")
        next
    end if
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
<input type="hidden" name="hInvoiceNo" value="<%=iSalInvoiceNo%>">
<input type="hidden" name="hFromDate" value="<%=sFinFromDate%>" />
<input type="hidden" name="hToDate" value="<%=sFinToDate%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%if sInvFrom="SI" then
			Response.Write "Credit Note For Sales Invoice"
		  elseif sInvFrom ="SR" then
			Response.Write "Credit Note For Sales Return"
		  end if
		%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
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
                                                    <td class="FieldCell" width="200">  <span id ="InoviceNo" class="DataOnly"><%=sInvoiceNo%></span></td>
                                                    <td class="FieldCellSub" width="100">Invoice Value</td>
                                                    <td class="FieldCellSub" width="145"> <span class="DataOnly"><%=FormatNumber(sInvValue,2,,,0)%> </span></td>
                                                        </tr>
                                                        
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Cr Note Against</td>
                                                    <td class="FieldCell" width="200">  
                                                    <Select name="SelCrAgain" class="FormElem" onChange="SetRetVal(this,'1')">
                                                    <Option Value="0">Select</Option>
                                                    <Option Value="Q" <%if sInvFrom="SR" then Response.Write "Selected" %> >Quantity</Option>
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
		sDiscPer =oNodEntry.Attributes.GetNamedItem("DisPer").value
		sDiscount=oNodEntry.Attributes.GetNamedItem("DisAmount").value
		dTotal=CDbl(dTotal)+CDbl(sAmount)

	if sInvFrom="SR" then		
		
		sQuery = "Select ItemCode,ClassificationCode,QuantityForReturn from Sal_T_SalesReturnDetail"&_
				 " where Salesreturnno in (Select SalesReturnNo from Sal_T_SalesReturnHeader where "&_
				 "SaleTransactionNo in(Select OtherApplnTransNo from Acc_T_CreatedVoucherHeader where CreatedTransNo ="& iInvNo &" ))"
			
		'Response.Write "iInvNo="& iInvNo
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source = sQuery
				.Open 
			end with 
			if not objRs.EOF then
				iitemCode = objrs(0)
				iClassCode = objrs(1)
				iSalRetQty = objrs(2)
			end if
			objRs.Close 
	end if ' 	if sInvFrom="SR" then		
%>

    <tr>
    <td class="ExcelSerial" align="center"><%=isno%></td>
    <%if sInvFrom="SR" then%>
		<td class="ExcelDisplayCell"><%=sDescription%>-<%=iitemCode%>-<%=iClassCode%></td>
    <%else%>
		<td class="ExcelDisplayCell"><%=sDescription%></td>
    <%end if%>
    <td class="ExcelDisplayCell" align="Right" id="tOldQty<%=isno%>"><%=FormatNumber(sQty,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" id="tOldRate<%=isno%>"><%=FormatNumber(sRate,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" id="tOldDis<%=isno%>"><%=FormatNumber(sDiscount,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount"  value="<%=FormatNumber(sAmount,2,,,0)%>" class="FormelemRead" size="13" readonly></td>
    <%if sInvFrom ="SR" then%>
		<td class="ExcelDisplayCell" align="Right" id="tQty<%=isno%>"><input type="text" style="text-align: Right" NAME="txtQty<%=isno%>" onBlur="setQty(this,'<%=iSno%>','Q')" value="<%=FormatNumber(iSalRetQty,3,,,0)%>" class="FormelemRead" size="13" readonly></td>
    <%else%>
		<td class="ExcelDisplayCell" align="Right" id="tQty<%=isno%>"><input type="text" style="text-align: Right" NAME="txtQty<%=isno%>" onBlur="setQty(this,'<%=iSno%>','Q')" value="<%=FormatNumber(sQty,3,,,0)%>" class="FormelemRead" size="13" readonly></td>
    <%end if%>
    <td class="ExcelDisplayCell" align="Right" id="tRate<%=isno%>"><input type="text" style="text-align: Right" NAME="txtRate<%=isno%>" onBlur="setQty(txtQty<%=iSno%>,'<%=iSno%>','R')" value="<%=FormatNumber(sRate,2,,,0)%>" class="FormelemRead" size="13" readonly></td>
    <input type="hidden" name="hRatePer" value="<%=sRatePer%>">
    </td>
    <td class="ExcelDisplayCell" align="Right" id="tDis<%=isno%>">
        <input type="hidden" name="hDisPer<%=iSNo%>" value="<%=sDiscPer%>">
    <input type="text" style="text-align: Right" NAME="txtDis<%=isno%>" value="<%=FormatNumber(sDiscount,2,,,0)%>" class="FormelemRead" size="13" onBlur="setQty(txtQty<%=iSno%>,'<%=iSno%>','R')" readonly>
    </td>
    

    <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount<%=iSno%>" onBlur="setTotal(this,'<%=iSno%>')" value="<%=FormatNumber(sAmount,2,,,0)%>" class="Formelem" size="13"></td>
	<!--td class="ExcelDisplayCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount"  value="<%=FormatNumber(sAmount,2,,,0)%>" class="FormelemRead" size="13"></td-->
        </tr>
<%
	next
%>


        <tr>
    <td align="center" ><Input type="hidden" name="hRowVal" value="<%=isno%>"></td>
   
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
		If iRndOff = 1 Then
			dTax = FormatNumber(Round(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,0),2,,,0)
		Else
			dTax = FormatNumber(oNodEntry.Attributes.GetNamedItem("TaxAmount").value,2,,,0)
		End If
%>
			<tr>
				<td align="center" colspan="1"></td>
				<td class="ExcelSerial" align="right" colspan="3"><%=sTaxName%>&nbsp;</td>
				<%if sTaxMode="P" then %>
				<td class="ExcelDisplayCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=FormatNumber(dTaxValue,2,,,0)%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="6" class="FormelemRead" readonly>&nbsp;%</td>
				<%else%>
				<td class="ExcelDisplayCell" align="right">
				<%
					if sTaxMode="K" then Response.Write "Per Pack"
					if sTaxMode="Q" then Response.Write "Per Qty"
				%>
				</td>
				<%end if%>
				<td class="ExcelDisplayCell" align="right">
				<input type="text" style="text-align: Right" NAME="txtTaxValue" value="<%=dTax%>"  size="11" class="FormelemRead" readonly></td>
				<td align="center" colspan="3"></td>
				<td class="ExcelInputCell" align="right">
				<%if sTaxMode="Q" or sTaxMode="K" then %>
					<input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  size="11" class="Formelem" onBlur="ReTotalCr()"></td>
				<%Else%>
					<input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  size="11" class="Formelem" onBlur="ReTotalCr()"></td>
				<%End IF %>
				
				    </tr>
				    
			<%
	next
	
	
oDOM.save server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")

%>

		

        <tr>
        <td align="center" colspan="1"></td>
    <td class="ExcelSerial" align="right" colspan="4"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelDisplayCell" align="right"> <input type="text" style="text-align: Right" NAME="txtFInvValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="FormelemRead"></td>
    <td align="right" colspan="3"></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="Formelem" readonly>
    
    </td>
        </tr>
        
        <tr>
        <td align="center" colspan="1"></td>
        <td class="ExcelSerial" align="right" colspan="8"><b>Credit Note Value&nbsp; </b></td>
    <!--td class="ExcelSerial" align="right" colspan="3"><b>Credit Note Value&nbsp; </b></td-->
    <td class="ExcelInputCell" align="right" id="tCrValue"> <input type="text" style="text-align: Right" NAME="txtCrNoteValue"  size="13" value="0.00" class="Formelem">
    
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
			
			<Textarea name="txtNarration" class="FormElem" cols="40" rows="4" maxlength="200"></Textarea>
			
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
