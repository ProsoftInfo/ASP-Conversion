<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNSalReturnEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 01, 2003
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
dim sDiscPer,dBasicTotal,dDisTotal,oNodtemp,sInvValue, iRndOff
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserId
Dim dPreCrValue,Temparr,sCallFrom ,sTotReturnQty

Dim sFinPeriod,sFromYr,sToYr,sTempYr,sFinFrm,sFinTo

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

sFinPeriod = Session("FinPeriod")
sTempYr = Split(sFinPeriod,":")
sFinFrm = Trim(sTempYr(0))
sFinTo = Trim(sTempYr(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

sInvoiceNo=Request.Form("selInvoiceNo")
sBookName=Request.Form("hBookName")
Temparr = Split(sInvoiceNo,":")
sInvoiceNo = Temparr(0)
dPreCrValue = Temparr(1)
'Response.write sInvoiceNo
iInvNo = sInvoiceNo
sCallFrom = Request.QueryString("hCallFrom")
sUserid = getUserID()

sQuery = "Select VoucherAmount From Acc_T_CreatedVoucherHeader Where CreatedTransNo "&_
		 "IN (Select CreatedTransNo From Acc_T_CreatedVoucherHeader Where BankInstrumentNo = '"&sInvoiceNo&"') "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	dPreCrValue = objRs(0)
End IF
objRs.Close
Dim sRetVal
sRetVal = GetVouchXML(sInvoiceNo)
oDOM.Load server.MapPath(sRetVal)
'oDOM.load server.MapPath("../xmldata/Voucher/"&sInvoiceNo&".xml")

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
dim dInvAmount
dInvAmount = sInvValue

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
Function SaveXML()

	Dim sExp,TempNode,sCheckVal,dOldInvValue,dNewInvValue,dPreCrValue
	dim iCrAccHead,sDesc,oDOMGJ,oDGjRoot,oDGjEntry,oDGjAcc,oDGjNarr,oDSubNode
	Dim oDOM,oNodRoot,oNodDeatils,oNodHeader,oNodEntry,oNodTaxRoot,objRs,newElem,newElem1
	Dim sPartyDet,sArrValue,sRetVal2,setFlag
	setFlag = false

	dOldInvValue = Trim(document.formname.hTotinvVal.value)
	dNewInvValue = Trim(document.formname.txtInvValue.value)
	dPreCrValue = Trim(document.formname.hPreCrValue.value)
	dOldInvValue = CDbl(dOldInvValue)
	dNewInvValue = CDbl(dNewInvValue)
'	dPreCrValue = CDbl(dPreCrValue)

	'dNewInvValue = Cdbl(dNewInvValue + dPreCrValue)

	IF dOldInvValue < dNewInvValue Then
		MsgBox "Returned Invoice Value Should be less than the Invoiced Value "
		document.formname.txtInvValue.focus
		Exit Function
	End IF
	set RootNode=TaxData.documentElement

	if Trim(document.formname.hCallfrom.value)="CR" then

			For Each oNodTemp in RootNode.childNodes
				if oNodTemp.nodeName="Details" then
				 	oNodTemp.Attributes.GetNamedItem("VouDate").value=document.formname.ctlDate.GetDate
				end if
			next

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

				Set newElem  = TaxData.createAttribute("CrTransNo")
				newElem.value = document.formname.hdTransNo.value
				TempNode.Item(0).setAttributeNode(newElem)

			End IF

			IF CheckFinDate() Then
				set objhttp = CreateObject("Microsoft.XMLHTTP")
				objhttp.Open "POST","XMLSave.asp?Mod=CN&Name=Voucher Entry", false
				objhttp.send TaxData.XMLDocument
				if objhttp.responseText <> "" then
					Msgbox(objhttp.responseText)
				else
					document.formname.B2.disabled = True
					document.formname.action="VouCNSalRetAdj.asp"
					document.formname.submit()
				end if
			End IF
	else ' if Trim(document.formname.hCallfrom.value)="CN" then

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
								oDGjEntry.setAttribute "Amount",document.formname.txtInvValue.Value
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
													NodDoc.setAttribute "AmtToAdjust", document.formname.txtinvoi.Value
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
							oDGjEntry.setAttribute "No",iCnt
							oDGjEntry.setAttribute "CRDR","D"
							oDGjEntry.setAttribute "Payto",oNodEntry.getAttribute("PayTo")
							oDGjEntry.setAttribute "Amount",document.formname.txtInvValue.Value
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
								oDGjNarr.Text = document.formname.txtNarration.value
								oDGjEntry.appendChild oDGjNarr
							next
						end if
					next
				end if
			next

			IF CheckFinDate() Then
					set objhttp = CreateObject("Microsoft.XMLHTTP")
					objhttp.Open "POST","XMLSave.asp?Mod=GJ&Name=Voucher Entry", false
					objhttp.send GJVoucher.XMLDocument
					if objhttp.responseText <> "" then
						Msgbox(objhttp.responseText)
					else
						document.formname.B2.disabled = True
						document.formname.action="VouGenerate.asp?hCallFrom=SR&hReturnQty="&document.formname.hTotReturnQty.value
						document.formname.submit()
					end if
			End IF
	end if ' if Trim(document.formname.hCallfrom.value)="CN" then
End Function

Function EnbApp(sObj)
	IF sObj.value = "Y" Then
		document.formname.selUserId.disabled = False
	Else
		document.formname.selUserId.selectedIndex = 0
		document.formname.selUserId.disabled = True
	End IF
End Function

Function SetDate()
	Dim sFromYr,sToYr
	sFromYr = document.formname.hFromYr.Value
	sToYr = document.formname.hToYr.Value
	sFromYr = "01/04/"&Trim(sFromYr)
	sToYr = "31/03/"&sToYr
	document.formname.ctlDate.setMinDate() = sFromYr
	document.formname.ctlDate.setMaxDate() = sToYr
End Function

Function CheckNoSer()
	Dim ObjHttp,sPassVal,sMon,sYear,sDate,sPeriod,sRetVal
	Set ObjHttp = CreateObject("MSXML2.XMLHTTP")
	IF Cstr(document.formname.hVouCode.Value) = "04" Then
		sPassVal = document.formname.selUnitId.Value 'Unit
		sPassVal = sPassVal&":"&document.formname.hVouCode.Value 'BookCode
		sPassVal = sPassVal&":"&document.formname.hCallFrm.Value 'Call For Created or Accounted
		sPassVal = sPassVal&":D"
		sPassVal = sPassVal&":"&document.formname.selBook.Value 'Book Number
	Else
		sPassVal = document.formname.hOrgid.Value 'Unit
		sPassVal = sPassVal&":"&document.formname.hVouCode.Value 'BookCode
		sPassVal = sPassVal&":"&document.formname.hCallFrm.Value 'Call For Created or Accounted
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
	'MsgBox sCurrMonYr &" " & sFinFrm &" " & sFinTo

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
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname">
<Input type="hidden" name="hdTransNo" value="<%=iInvNo%>">
<Input type="hidden" name="hPreCrValue" value="<%=dPreCrValue%>">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hBookCode" value="<%=Request.Form("selBook")%>">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">
<input type="hidden" name="hBookName" value="<%=Request.Form("hBookName")%>">

<input type="hidden" name="txtNarration" value="">
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
		<td align="center" class=PageTitle height="20"><p align="center">Sales Return Credit
          Note 		</td>
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
								<td class="TabCell" valign="bottom" align="center" width="90">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Adjustments
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
                            </tr>
                            <tr>
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
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="100">Unit </td>
                                                    <td class="FieldCell" width="200">  <span class="DataOnly"><%=sOrgName%></span></td>
                                                    <td class="FieldCellSub" width="100"><p align="left">Date</p></td>
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
                                                    <td class="FieldCellSub" width="145"> <span class="DataOnly"><%=sInvValue%></span></td>
                                                        </tr>

                                                        <tr>
															<td class="FieldCellSub" width="100">Credit Note Value</td>
															<td class="FieldCell" width="200">  <span class="DataOnly"><%=FormatNumber(dInvAmount,2,,,0)%></span></td>
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
                <div class="frmBody" id="frm2" style="width: 575; height:320;">
            <table border="0" cellspacing="1" class="ExcelTable" width="100%">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
    <td class="ExcelHeaderCell" align="center" colspan="3">Invoice</td>
    <td class="ExcelHeaderCell" align="center" colspan="2">Returned</td>

        </tr>
        <tr>
    <td class="ExcelHeaderCell" align="center" width="75">Qty</td>
    <td class="ExcelHeaderCell" align="center" width="75">Value</td>
    <td class="ExcelHeaderCell" align="center" width="75">Discount</td>
    <td class="ExcelHeaderCell" align="center">Qty</td>
    <td class="ExcelHeaderCell" align="center">Amount</td>

        </tr>
<%
	For Each oNodEntry in oNodDeatils.childNodes
		iSno=oNodEntry.Attributes.GetNamedItem("No").value
		sDescription=oNodEntry.Attributes.GetNamedItem("PayTo").value
		sAmount=oNodEntry.Attributes.GetNamedItem("Amount").value
		sQty=oNodEntry.Attributes.GetNamedItem("Qty").value
		sValue=oNodEntry.Attributes.GetNamedItem("ActValue").value
		sDiscount=oNodEntry.Attributes.GetNamedItem("DisAmount").value
		oNodEntry.Attributes.GetNamedItem("Rate").value=CDbl(oNodEntry.Attributes.GetNamedItem("Amount").value)/CDbl(sQty)
		dTotal=CDbl(dTotal)+CDbl(sAmount)

%>

	<Input type="hidden" name=hQty<%=iSno%> value="<%=sQty%>">
	<Input type="hidden" name=hInvVal<%=iSno%> value="<%=sValue%>">

    <tr>
    <td class="ExcelSerial" align="center"><%=isno%></td>
    <td class="ExcelDisplayCell"><%=sDescription%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sQty,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sValue,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="75"><%=FormatNumber(sDiscount,2,,,0)%></td>
    <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtqty<%=iSno%>1" onBlur="setQty(this,'<%=iSno%>')" value="<%=FormatNumber(sQty,3,,,0)%>" class="Formelem" size="13"></td>

    <td class="ExcelInputCell" align="Right"><input type="text" style="text-align: Right" NAME="txtAmount<%=iSno%>" onBlur="setTotal(this,'<%=iSno%>')" value="<%=FormatNumber(sAmount,2,,,0)%>" class="Formelem" size="13"></td>

        </tr>
<%
	sTotReturnQty = cdbl(sTotReturnQty) + cdbl(sQty)
	next
%>


        <tr>
    <td align="center" colspan="2"></td>
    <input type="hidden" name="hTotReturnQty" value="<%=sTotReturnQty%>">
    <td class="ExcelSerial" align="center"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
     <td align="right" colspan="2"></td>
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
				<td align="center" colspan="2"></td>
				<td class="ExcelSerial" align="center" colspan="3"><%=sTaxName%>&nbsp;</td>
				<%if sTaxMode="P" then %>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxValue%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="6" class="Formelem">&nbsp;%</td>
				<%else%>
				<td class="ExcelDisplayCell" align="right">
				<%
					if sTaxMode="K" then Response.Write "Per Pack"
					if sTaxMode="Q" then Response.Write "Per Qty"
				%>
				</td>
				<%end if%>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  onBlur="setTaxAmount('<%=sCatCode%>','<%=sTaxCode%>',this)"size="11" class="Formelem"></td>
				    </tr>
			<%
	next
oDOM.save server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")

%>

		<input type="hidden" name="hTotinvVal" value="<%=dInvAmount%>">
        <tr>
        <td align="center" colspan="2"></td>
    <td class="ExcelSerial" align="right" colspan="4"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  size="13" value="<%=FormatNumber(dInvAmount,2,,,0)%>" class="Formelem">
    </td>
        </tr>
        <tr>
        </tr>
        <tr>
			<td align="left" class="FieldCell" colspan="2" valign="Top">
				Approval
			</td>
			<td align="center" class="FieldCellSub" colspan="2">
			<Input type="radio" name="optApprove" checked value="Y" onClick="EnbApp(this)"> Yes &nbsp;&nbsp;&nbsp;
			<Input type="radio" name="optApprove" value="N" onClick="EnbApp(this)"> No &nbsp;&nbsp;&nbsp;
			</td>
        </tr>
        <tr>
			<td align="left" class="FieldCell" colspan="2" valign="Top">
				Immediate Approver
			</td>
			<td align="center" class="FieldCellSub" colspan="2">
			<select size="1" name="selUserId" class="FormElem">
              <option value="I">Immediate Approver</option>
                <%=populateEmployeeWithVal(sUserId)%>
                    </select>
			</td>
        </tr>
            </table>
                </div>
                </td>
                <td></td>
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
