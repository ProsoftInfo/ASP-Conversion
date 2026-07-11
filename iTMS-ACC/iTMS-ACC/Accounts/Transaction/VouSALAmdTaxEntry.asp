<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouSALAmdTaxEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February  18, 2003
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
<!--#include file="../../include/populate.asp"-->
<%
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal
dim sSalType,sOrgId,sorgName,sQuery,sPartyName,sRefernceNo
dim sDiscPer,dBasicTotal,dDisTotal, iRndOff,iTransno,sFlag,sFrmOth
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sOthAppChk
dim sInvoiceNo,sAppTy,sStr,TempNode,Objrs2,sUserId,sRndChk,pTransNo,sAmdTy,sDisable
Dim iBookNo,iBookAcc,iOldBookNo,sChgTy
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")
sAppTy = Request.Form("OptApproval")
pTransNo = Request("hTransNo")
sFlag = Request("hFlag")
sUserId = getUserID()
sAmdTy = Request("hAmdTy")
sFrmOth = Request("hFrmOthApp")

IF CStr(sFrmOth) = "" Then
	sFrmOth = "0"
End IF

'Response.Write sFrmOth

IF CStr(Trim(sFrmOth)) = "0" Then
	sDisable = ""
Else
	sDisable = "disabled"
End IF

'Response.Write pTransNo

IF CStr(sAmdTy) = "A" Then
	sFlag = True
End IF
oDOM.load server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")

set oNodRoot=oDOM.documentElement
sStr = "//Voucher"
Set TempNode = oNodRoot.selectNodes(sStr)
IF TempNode.length <> 0 Then
	TempNode.Item(0).Attributes.Item(0).nodeValue = pTransNo
End IF
iTransno = pTransNo

for each oNodDeatils in oNodRoot.childNodes
	if oNodDeatils.nodeName="Header" then
		for Each oNodEntry in  oNodDeatils.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sorgName=oNodEntry.text
			end if
			if oNodEntry.nodeName="Book" then
				iBookNo = oNodEntry.Attributes.Item(0).nodeValue
				iBookAcc = oNodEntry.Attributes.Item(1).nodeValue
			end if
			if oNodEntry.nodeName="SalesType" then
				sSalType=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sInvoiceNo=oNodEntry.Attributes.Item(0).nodeValue &"&nbsp; Dt:"& oNodEntry.Attributes.Item(1).nodeValue
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="RefNo" then
				sRefernceNo=oNodEntry.Text
			end if

		next
	end if
	if oNodDeatils.nodeName="Details" then exit for
next

Dim dTotNewBasVal,iCtr,sExp
dTotNewBasVal = 0
sExp = "//Entry"
Set TempNode = oNodRoot.selectNodes(sExp)
For iCtr = 0 To TempNode.Length - 1
	dTotNewBasVal = Cdbl(TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value) + Cdbl(dTotNewBasVal)
Next

sQuery = "Select RoundOff From App_R_ORGNTaxAccountHead Where InvoiceType = "&sSalType&" "&_
		 "and TaxCode is NULL and OUDefinitionID = '"&sOrgId&"' "

objRs.Open sQuery,con
IF Not objRs.EOF Then
	sRndChk = objRs(0)
End IF
objRs.Close

'IF sRndChk = 0 Not to Roundoff the Invoice Value 1 It has to get Rounded off

sQuery = "Select isNull(FromApplication,0),BookNumber From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "& iTransNo
objRs.Open sQuery,con
IF Not objRs.EOF Then
	sOthAppChk = objRs(0)
	iOldBookNo = objRs("BookNumber")
End IF
objRs.Close

IF CStr(iOldBookNo) <> CStr(iBookNo) Then
	sChgTy = "Y"
Else
	sChgTy = "N"
End IF

'Response.Write sChgTy



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML id="TaxData" src="<%="../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"%>"></XML>
<script SRC="../scripts/VouSalesPurchase.js"></SCRIPT>
<script SRC="../../scripts/TaxEntryCompat.js"></SCRIPT>
<script>
function CheckSubmit() {
	ITMSTaxEntryCompat.checkSubmit({
		invoiceNodeName: "SaleInvoice",
		saveUrl: "XMLSave.asp?Name=Voucher AMD&Mod=SAL",
		skipZeroTax: true,
		enableTaxSelects: true,
		beforeSubmit: function (form) {
			if (ITMSTaxEntryCompat.trim(form.hFlag.value) === "True") {
				form.action = "VouSALAmdAccAdvance.asp";
			} else {
				form.action = "VouSALAmdAdvance.asp";
			}
		}
	});
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="VouSALAmdAdvance.asp">
<Input type="hidden" name="hApprover" value="<%=sAppTy%>">
<Input type="hidden" name="hRndChk" value="<%=sRndChk%>">
<input type="hidden" name="hTransNo" value="<%=pTransNo%>">
<input type="hidden" name="hFlag" value="<%=sFlag%>">
<input type="hidden" name="hAmdTy" value="<%=sAmdTy%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Voucher 		</td>
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
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								  	<tr>
									  	<td align="center">Invoice Details</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="75">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<td align="center">Advance</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Voucher</td>
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
                            <tr>
                            <td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                            <!--tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: pointer" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">ª</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Daywise Balance"><font size="3" face="Webdings">¦</font>                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Voucher History">
                    <font size="4" face="Webdings">¨</font>
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
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr-->
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0"  cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <!--<tr>
                                                    <td class="FieldCellSub" width="75">Unit</td>
                                                    <td class="FieldCell" colspan="3">  <span class="DataOnly"><%=sorgName%>&nbsp;</span></td>

                                                        </tr>-->
                                                        <tr>
															<td class="FieldCellSub" width="75">Party Name</td>
															<td class="FieldCell" colspan="3">  <span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
														</tr>
														<tr>
															<td class="FieldCellSub" width="110"><p align="left">Invoice No</p></td>
															<td class="FieldCell" colspan="3"> <span class="DataOnly"><%=sInvoiceNo%>&nbsp;</span> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
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
                <div class="frmBody" id="frm2" style="width: 600; height:280;">
            <table border="0" cellspacing="1" class="ExcelTable" width="574">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Invoice<br>
    Quantity</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Invoice<br>
    Rate</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Rate<br>
    Per</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Basic<br>
    Value</td>
    <td class="ExcelHeaderCell" align="center" colspan="2">Discount</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Nett<br>
    Basic</td>
        </tr>
        <tr>
    <td class="ExcelHeaderCell" align="center" width="25">%</td>
    <td class="ExcelHeaderCell" align="center" width="60">Value</td>
        </tr>
<%
	Dim iLoop,dRndOff,dRatePer
	iLoop = 1
	For Each oNodEntry in oNodDeatils.childNodes
		iSno=oNodEntry.Attributes.GetNamedItem("No").value
		sDescription=oNodEntry.Attributes.GetNamedItem("PayTo").value
		sAmount=oNodEntry.Attributes.GetNamedItem("Amount").value
		sRate=oNodEntry.Attributes.GetNamedItem("Rate").value
		sQty=oNodEntry.Attributes.GetNamedItem("Qty").value& "&nbsp;"&oNodEntry.Attributes.GetNamedItem("UOMValue").value
		sValue=oNodEntry.Attributes.GetNamedItem("ActValue").value
		sDiscPer=oNodEntry.Attributes.GetNamedItem("DisPer").value
		sDiscount=oNodEntry.Attributes.GetNamedItem("DisAmount").value
		IF Cstr(sOthAppChk) = "0" Then
			dRndOff = oNodEntry.Attributes.GetNamedItem("RndOff").value
		End IF
		dRatePer = oNodEntry.Attributes.GetNamedItem("RatePer").value

		sAmount = CDbl(oNodEntry.Attributes.GetNamedItem("Qty").value) * (CDbl(sRate) /CDbl(dRatePer))
		sAmount = CDbl(sAmount) - CDbl(sDiscount)
		sAmount = CDbl(sAmount) + CDbl(dRndOff)
		sAmount = FormatNumber(sAmount,2,,,0)

		IF Len(Trim(sDisable)) = 0 Then
			oNodEntry.Attributes.GetNamedItem("Amount").value = sAmount
		Else
			sAmount = sValue
			oNodEntry.Attributes.GetNamedItem("Amount").value = sAmount
		End IF


		dTotal=CDbl(dTotal)+CDbl(sAmount)
		dBasicTotal=CDbl(dBasicTotal)+CDbl(sValue)
		dDisTotal=CDbl(dDisTotal)+CDbl(sDiscount)


%>
    <tr>
    <td class="ExcelSerial" align="center"><%=iLoop%></td>
    <td class="ExcelDisplayCell"><%=sDescription%></td>
    <td class="ExcelDisplayCell" align="Left" width="60"><%=sQty%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=sRate%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(dRatePer,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sValue,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="25"><%=FormatNumber(sDiscPer ,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber( sDiscount,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sAmount,2,,,0)%></td>
        </tr>
<%
	iLoop = iLoop + 1
	next
%>

        <tr>
    <td align="center" colspan="4" class="ExcelSerial"><b>Account Head</b></td>
    <td class="ExcelSerial" align="center"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dBasicTotal,2,,,0)%></b></td>
    <td class="FieldCell" align="center" width="25">   </td>
    <td class="ExcelDisplayCell" align="right" width="60"><b><%=FormatNumber(dDisTotal,2,,,0)%></b></td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
        </tr>
        <input type="Hidden" name="hBasicValue" value="<%=dTotNewBasVal%>">
        <input type="Hidden" name="hDisValue" value="<%=dDisTotal%>">
        <input type="Hidden" name="hAmount" value="<%=dTotal%>">
<%
oNodDeatils.Attributes.GetNamedItem("BasicValue").value=dTotNewBasVal
oNodDeatils.Attributes.GetNamedItem("Discount").value=dDisTotal
oNodDeatils.Attributes.GetNamedItem("ActualValue").value=dTotal




dim sAccHead,oNodTaxRoot,sTempTaxAmt
Set oNodTaxRoot = oDOM.createElement("TaxDetails")
oNodTaxRoot.setAttribute "InvoiceVlaue","0"
oNodTaxRoot.setAttribute "Basicvalue",dBasicTotal
oNodTaxRoot.setAttribute "NettValue",dTotal
oNodTaxRoot.setAttribute "RoundOffValue",dTotal

'sQuery = "select TaxShortName,TaxCategoryCode,TaxCode,ComputationMode,isnull(SumOfFields,''),isnull(FlatAmount,0),"&_
'	    " AccountHead, ROUNDOFF from VwSalesTaxDetails where ComputationMode is not null and  OUDefinitionID='"&sOrgId&"' and AccountTaxAccHead=1 and InvoiceType="&sSalType&" order by TaxHierarchy"

IF CStr(sChgTy) = "N" Then
	sQuery = "SELECT V.TAXSHORTNAME,V.TAXCATEGORYCODE,V.TAXCODE,V.COMPUTATIONMODE,ISNULL(V.SUMOFFIELDS,''), "&_
			 "ISNULL(V.FLATAMOUNT,0), C.ACCOUNTHEAD, V.ROUNDOFF FROM VWSALESTAXDETAILS V,  "&_
			 "Acc_T_CreatedVoucherTaxDet C WHERE V.COMPUTATIONMODE IS NOT NULL AND  "&_
			 "V.OUDEFINITIONID='"&sOrgId&"' AND V.ACCOUNTTAXACCHEAD=1 AND V.TAXCATEGORYCODE =  "&_
			 "C.TAXCATEGORYCODE and V.TaxCode = C.TaxCode and C.CreatedTransNo = "&iTransno&" And "&_
			 "V.INVOICETYPE="&sSalType&" ORDER BY V.TAXHIERARCHY  "
Else
	sQuery = "Select TaxShortName,TaxCategoryCode,TaxCode,ComputationMode,isnull(SumOfFields,''),isnull(FlatAmount,0),"&_
		     " AccountHead, ROUNDOFF from VwSalesTaxDetails where ComputationMode is not null and  OUDefinitionID='"&sOrgId&"' and AccountTaxAccHead=1 and InvoiceType="&sSalType&" order by TaxHierarchy"
End IF




'Response.Write sQuery

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing


do while not objRs.EOF

	sTaxName=objRs(0)
	sCatCode=objRs(1)
	sTaxCode=objRs(2)
	sTaxMode=objRs(3)
	sFormula=objRs(4)
	dTaxValue=objRs(5)
	sAccHead=objRs(6)
	iRndOff = objRs(7)

	IF CStr(sChgTy) = "N" Then
		sQuery = "Select isnull(TaxPercentage,0),isNull(TaxAmount,0) From Acc_T_CreatedVoucherTaxDet Where "&_
				 "CreatedTransNo = "&iTransno&" and TaxCode = "&sTaxCode&" and TaxCategoryCode = "&sCatCode&" "

		'Response.Write sQuery &"<br>"
		with objRs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs2.ActiveConnection = nothing
		IF Not objRs2.EOF Then
			dTaxValue = objRs2(0)
			sTempTaxAmt = Objrs2(1)
		End IF
		objRs2.Close
	Else
		sTempTaxAmt = 0
	End IF

	IF CStr(sTaxMode) = "F" and CStr(sChgTy) = "N" Then
		dTaxValue = sTempTaxAmt
	End IF

	Set newElem = oDOM.createElement("Tax")
	newElem.setAttribute "CatCode",sCatCode
	newElem.setAttribute "TaxCode",sTaxCode
	newElem.setAttribute "TaxMode",sTaxMode
	newElem.setAttribute "TaxFormula",sFormula
	newElem.setAttribute "TaxValue",CStr(dTaxValue)
	newElem.setAttribute "TaxAmount", sTempTaxAmt
	newElem.setAttribute "AccHead",sAccHead
	newElem.setAttribute "RoundOff",iRndOff
	newElem.Text= sTaxName
	oNodTaxRoot.appendChild newElem

	objRs.MoveNext
loop
objRs.Close

dim dInvAmount,iRoundedInvvalue,iRoundedoff

	dInvAmount=dTotal
	For Each oNodEntry in oNodTaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.GetNamedItem("CatCode").value
		sTaxCode=oNodEntry.Attributes.GetNamedItem("TaxCode").value
		sTaxMode=oNodEntry.Attributes.GetNamedItem("TaxMode").value
		sFormula=oNodEntry.Attributes.GetNamedItem("TaxFormula").value
		dTaxValue = oNodEntry.Attributes.GetNamedItem("TaxValue").value
		iRndOff = oNodEntry.Attributes.GetNamedItem("RoundOff").value
		sAccHead = oNodEntry.Attributes.GetNamedItem("AccHead").value
		dTax = oNodEntry.Attributes.GetNamedItem("TaxAmount").value
		sTaxName=oNodEntry.Text


		if sTaxMode="P" and CStr(sChgTy) = "Y" then
			dTax=CalculateTax(oNodTaxRoot,sFormula,dBasicTotal,dTotal,dTaxValue)
		'else
		'	dTax=dTaxValue
		end if

		If iRndOff = 1 and CStr(sChgTy) = "Y" Then
			dTax = Round(dTax,0)
			dTax = FormatNumber(dTax,2,,,0)
		'Else
		'	dTax=FormatNumber(dTax,2,,,0)
		End If

		dInvAmount=dInvAmount+CDbl(dTax)
		dTax=FormatNumber(dTax,2,,,0)
		dInvAmount=FormatNumber(dInvAmount,2,,,0)
		dTaxValue=FormatNumber(dTaxValue,2,,,0)
		'check
		'Response.Write "dTaxValue="&dTaxValue&"<BR>"
		'oNodEntry.Attributes.GetNamedItem("TaxAmount").value=dTax
			%>
			<tr>
				<td align="center" colspan="3" class="ExcelSerial">
				<Select name="SelAccHead<%=sCatCode%><%=sTaxCode%>" class="FormElem" <%=sDisable%>>
				<Option Value="0" Selected>ADD WITH ITEM VALUE</Option>
					<%=PopulateTaxAccHead()%>
				</Select>
				</td>
				<td class="ExcelSerial" align="center" colspan="4"><%=sTaxName%>&nbsp;</td>
				<% if sTaxMode="P" then %>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxValue%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="7" class="Formelem" <%=sDisable%>>&nbsp;%</td>
				<%else%>
				<td class="ExcelDisplayCell" align="right">
				<%
					if sTaxMode="K" then Response.Write "Per Pack"
					if sTaxMode="Q" then Response.Write "Per Qty"

				%>
				</td>
				<%end if%>
				<% IF CStr(sTaxMode) = "F" Then %>
					<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  onBlur="setTaxAmount('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="13" size="15" class="Formelem" <%=sDisable%>></td>
				<%Else%>
					<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTax%>"  onBlur="setTaxAmount('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="13" size="15" class="Formelem" <%=sDisable%>></td>
				<%End IF %>
				    </tr>
			<%
	next

	Dim OldTaxNode,iChkCtr
	sExp = "//TaxDetails"
	Set OldTaxNode = oNodRoot.selectNodes(sExp)
	IF OldTaxNode.length <> 0 Then
		'OldTaxNode.removeAll
		For iChkCtr = 0 To OldTaxNode.length - 1
					oNodRoot.RemoveChild OldTaxNode.Item(iChkCtr)
		Next
	End IF

	IF CStr(sRndChk) = "0" Then 'if No Roundoff is Needed Then
		iRoundedInvvalue = dInvamount
	Else
		iRoundedInvvalue = round(dInvamount)
	End IF



	iRoundedoff = Round(cdbl(iRoundedInvvalue) - cdbl(dInvamount),2)
	oNodTaxRoot.Attributes.GetNamedItem("NettValue").value=iRoundedInvvalue
	oNodTaxRoot.Attributes.GetNamedItem("RoundOffValue").value=iRoundedoff
	oNodRoot.appendChild oNodTaxRoot

	oNodTaxRoot.Attributes.GetNamedItem("InvoiceVlaue").value=dInvAmount
	oNodRoot.appendChild oNodTaxRoot
oDOM.save server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")

%>
  <tr>

    <td class="ExcelSerial" align="center" colspan="8">
    <p align="right"><b>Rounded Off&nbsp; </b></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtroundoff"  size="11" value="<%=iRoundedoff%>"  class="Formelem" readonly <%=sDisable%>>
        </tr>

        <tr>

    <td class="ExcelSerial" align="right" colspan="8"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  Maxlength="13" size="15" value="<%=FormatNumber(iRoundedInvvalue,2,,,0)%>" class="Formelem" <%=sDisable%>>
    </td>
        </tr>

            </table>
                </div>
                </td>
                <td></td>
                            </tr>
                            <tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td>
												 <% IF CStr(sAppTy) = "Y" Then %>
												 <tr>
												     <td class="FieldCell" >Immediate Approver</td>
												     <td class="FieldCellSub">
												     <select size="1" name="selUserId" class="FormElem">
												 <option value="I">Immediate Approver</option>
												 <%=populateEmployeeWithVal(sUserId)%>
												     </select></td>
												         </tr>
												<%End IF %>
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
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
                                                                <input type="button" value="Next" name="B2" class="ActionButton" onClick="CheckSubmit()" >
                                                                <input type="button" value="Cancel" name="B6" class="ActionButton" >
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
<%
	Function PopulateTaxAccHead()
		sQuery = "Select Left(AccountDescription,45),AccountHead From VWACCHEADFORSALESAPP  "&_
				 "Where OUDefinitionID = '"&sOrgId&"' Order By 1"

		With objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End WIth
		Set objRs.ActiveConnection = Nothing
		Do While Not objRs.EOF
			IF CStr(sAccHead) = CStr(objRs(1)) Then
%>
				<option value="<%=objRs(1)%>" Selected><%=objRs(0)%></option>
<%			Else		%>
				<option value="<%=objRs(1)%>"><%=objRs(0)%></option>
<%
			End IF

			objRs.MoveNext
		Loop
		objRs.Close
	End Function

%>
