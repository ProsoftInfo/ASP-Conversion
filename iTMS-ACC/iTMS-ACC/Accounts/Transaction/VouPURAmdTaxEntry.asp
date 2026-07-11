<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouPURAmdTaxEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	OCT 28, 2004
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
dim sSalType,sOrgId,sQuery,sPartyName,sInvoiceNo,dRndOff
dim sDiscPer,dBasicTotal,dDisTotal,sInvoiceDt,sAppValue
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue,sUserID
Dim sInvRndChk,Objrs2,iCrTransNo,sFormVal,sAmdTy,iTransNo

sUserID = getUserID()

sAppValue = Request.Form("optApprove")
sFormVal = Request.Form("hFormVal")
sAmdTy = Request("sAmdType")
iTransNo = Request("hTransNo")



' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")
oDOM.load server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")
'oDOM.load server.MapPath("../temp/transaction/Pur_Voucher_xml_"&Session.SessionID&".xml")
set oNodRoot=oDOM.documentElement

iCrTransNo = oNodRoot.Attributes.Item(0).nodeValue

'Response.write iCrTransNo

for each oNodDeatils in oNodRoot.childNodes
	if oNodDeatils.nodeName="Header" then
		for Each oNodEntry in  oNodDeatils.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="PurchaseType" then
				sSalType=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="PurInvoice" then
				sInvoiceNo=oNodEntry.Attributes.getNamedItem("PurInvNo").Value
				sInvoiceDt=oNodEntry.Attributes.getNamedItem("PurInvDate").Value
			end if

		next
	end if

 'Added by Maheswari on 23rd May
	IF trim(oNodDeatils.NodeName) = "TaxDetails" then
		dRndOff = oNodDeatils.Attributes.GetNamedItem("RoundOffValue").value
	'	Response.Write "<p>Chk="&dRndOff &"<BR>"
	End IF
next
for each oNodDeatils in oNodRoot.childNodes
	if oNodDeatils.nodeName="Details" then exit for
next
sQuery = "Select Top 1 InvRoundOff From APP_R_PurchaseOrgnTaxAccHead  "&_
		 "Where OUDefinitionID = '"&sOrgId&"' and PurChaseType = "&Trim(sSalType)&" "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sInvRndChk = objRs(0)
End IF
objRs.Close


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML id="TaxData" src="<%="../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml"%>"></XML>

<script SRC="../scripts/VouPurchase.js"></SCRIPT>
<script SRC="../../scripts/cancel.js"></SCRIPT>
<script SRC="../../scripts/TaxEntryCompat.js"></SCRIPT>
<script>
function CheckSubmit() {
	ITMSTaxEntryCompat.checkSubmit({
		invoiceNodeName: "PurInvoice",
		saveUrl: "XMLSave.asp?Mod=PUR&Name=Voucher Amd",
		skipZeroTax: true,
		disableButton: "B2"
	});
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="popTax()">
<form method="POST" name="formname" action="VouPURAmdAdvance.asp">
<Input type="hidden" name="hApprover" value="<%=sAppValue%>">
<Input type="hidden" name="hRndChk" value="<%=Trim(sInvRndChk)%>">
<input type="hidden" name="hFormVal" value="<%=sFormVal%>">
<input type="hidden" name="hAmdType" value="<%=sAmdTy%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Purchase Voucher Amendment	</td>
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
								<!--<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>-->
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
                    <span style="cursor: pointer" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">?</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Daywise Balance"><font size="3" face="Webdings">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Voucher History">
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
                                                            <table border="0" cellspacing="0" cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="75">Party Name</td>
                                                    <td class="FieldCell" colspan="3">  <span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
                                                    </tr>
                                                    <tr>
                                                    <td class="FieldCellSub" width="75">Invoice No</td>
                                                    <td class="FieldCell" width="200">  <span class="DataOnly"><%=sInvoiceNo%>&nbsp;</span></td>

                                                    <td class="FieldCellSub" width="110"><p align="left">Voucher Date</p></td>
                                                    <td class="FieldCellSub" width="145"> <span class="DataOnly"><%=sInvoiceDt%>&nbsp;</span> </td>
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


	Dim iRow
	iRow = 1

	For Each oNodEntry in oNodDeatils.childNodes

		iSno=oNodEntry.Attributes.GetNamedItem("No").value
		sDescription=oNodEntry.Attributes.GetNamedItem("PayTo").value
		sAmount=oNodEntry.Attributes.GetNamedItem("Amount").value
		sRate=oNodEntry.Attributes.GetNamedItem("Rate").value
		sQty=oNodEntry.Attributes.GetNamedItem("Qty").value& "&nbsp;"&oNodEntry.Attributes.GetNamedItem("UOMValue").value
		sValue=oNodEntry.Attributes.GetNamedItem("ActValue").value
		sDiscPer=oNodEntry.Attributes.GetNamedItem("DisPer").value
		sDiscount=oNodEntry.Attributes.GetNamedItem("DisAmount").value

		sAmount = Cdbl(sValue) - Cdbl(sDiscount)

		dTotal=CDbl(dTotal)+CDbl(sAmount)
		dBasicTotal=CDbl(dBasicTotal)+CDbl(sValue)
		dDisTotal=CDbl(dDisTotal)+CDbl(sDiscount)

%>
    <tr>
    <td class="ExcelSerial" align="center"><%=iRow%></td>
    <td class="ExcelDisplayCell"><%=sDescription%></td>
    <td class="ExcelDisplayCell" align="Left" width="60"><%=sQty%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sRate,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sValue,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="25"><%=FormatNumber(sDiscPer ,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber( sDiscount,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sAmount,2,,,0)%></td>
        </tr>
<%
		iRow = iRow + 1
	next
%>

        <tr>

    <td class="ExcelSerial" align="center" colspan="4"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dBasicTotal,2,,,0)%></b></td>
    <td class="FieldCell" align="center" width="25">    </td>
    <td class="ExcelDisplayCell" align="right" width="60"><b><%=FormatNumber(dDisTotal,2,,,0)%></b></td>
    <td class="ExcelDisplayCell" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
        </tr>
        <input type="Hidden" name="hBasicValue" value="<%=dBasicTotal%>">
        <input type="Hidden" name="hDisValue" value="<%=dDisTotal%>">
        <input type="Hidden" name="hAmount" value="<%=dTotal%>">
<%

oNodDeatils.Attributes.GetNamedItem("BasicValue").value=dBasicTotal
oNodDeatils.Attributes.GetNamedItem("Discount").value=dDisTotal
oNodDeatils.Attributes.GetNamedItem("ActualValue").value=dTotal

dim sAccHead,oNodTaxRoot,iRndOff,iRoundedInvvalue,iRoundedoff,dTaxPreAmt
Set oNodTaxRoot = oDOM.createElement("TaxDetails")
oNodTaxRoot.setAttribute "InvoiceVlaue","0"
oNodTaxRoot.setAttribute "Basicvalue",dBasicTotal
oNodTaxRoot.setAttribute "NettValue",dTotal
oNodTaxRoot.setAttribute "RoundOffValue",dRndOff

sQuery="select TaxShortName,TaxCategoryCode,TaxCode,ComputationMode,isnull(SumOfFields,''),isnull(FlatAmount,0),"&_
			" AccountHead,isNull(RoundOff,0) from VwPurchaseTaxDetails where ComputationMode is not null and  OUDefinitionID='"&sOrgId&"' and PurchaseType="&sSalType&" order by TaxHierarchy"
'Response.write sQUery &"<br>"
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

	sQuery = "Select isNull(TaxPercentage,0),TaxAmount,AccountHead From Acc_T_CreatedVoucherTaxDet Where  "&_
			 "CreatedTransNo = "&iCrTransNo&" and TaxCode = "&sTaxCode&" and TaxCategoryCode = "&sCatCode&"  "

	 'Response.write sQUery &"<br>"
	Objrs2.open sQuery,Con
	IF Not Objrs2.Eof Then
		dTaxValue = Trim(Objrs2(0))
		dTaxPreAmt = Trim(Objrs2(1))
		sAccHead = Trim(Objrs2(2))
	Else
		dTaxPreAmt = 0
		sAccHead = 0
	End IF
	Objrs2.Close

'Response.write dTaxPreAmt &"<br>"

	Set newElem = oDOM.createElement("Tax")
	newElem.setAttribute "CatCode",sCatCode
	newElem.setAttribute "TaxCode",sTaxCode
	newElem.setAttribute "TaxMode",sTaxMode
	newElem.setAttribute "TaxFormula",sFormula
	newElem.setAttribute "TaxValue",CStr(dTaxValue)
	newElem.setAttribute "TaxAmount",dTaxPreAmt
	newElem.setAttribute "AccHead",sAccHead
	newElem.setAttribute "ItemValue","0"
	newElem.setAttribute "RoundOff",iRndOff
	newElem.Text = sTaxName
	oNodTaxRoot.appendChild newElem

	objRs.MoveNext
loop
objRs.Close

dim dInvAmount,dTaxAmt
	dInvAmount=dTotal

	For Each oNodEntry in oNodTaxRoot.childNodes
	'Response.Write oNodEntry.nodename
		sCatCode=oNodEntry.Attributes.GetNamedItem("CatCode").value
		sTaxCode=oNodEntry.Attributes.GetNamedItem("TaxCode").value
		sTaxMode=oNodEntry.Attributes.GetNamedItem("TaxMode").value
		sFormula=oNodEntry.Attributes.GetNamedItem("TaxFormula").value
		dTaxValue=oNodEntry.Attributes.GetNamedItem("TaxValue").value
		dTaxAmt =oNodEntry.Attributes.GetNamedItem("TaxAmount").value
		iRndOff = oNodEntry.Attributes.GetNamedItem("RoundOff").value
		sAccHead = oNodEntry.Attributes.GetNamedItem("AccHead").value

		sTaxName=oNodEntry.Text

		if sTaxMode="P" then
			dTax=CalculateTax(oNodTaxRoot,sFormula,dBasicTotal,dTotal,dTaxValue)
		else
			dTax=dTaxValue
		end if

		If iRndOff = 1 Then
			dTax = Round(dTax,0)
			dTax = FormatNumber(dTax,2,,,0)
		Else
			dTax=FormatNumber(dTax,2,,,0)
		End If

		dTaxValue=FormatNumber(dTaxValue,2,,,0)

		dInvAmount=dInvAmount+CDbl(dTax)
		dTax=FormatNumber(dTax,2,,,0)
		dInvAmount=FormatNumber(dInvAmount,2,,,0)

		oNodEntry.Attributes.GetNamedItem("TaxAmount").value=dTax


			%>
			<tr>
				<td align="center" colspan="3" class="ExcelSerial">
				<Select name="SelAccHead<%=sCatCode%><%=sTaxCode%>" class="FormElem">
				<Option Value="0" Selected>ADD WITH ITEM VALUE</Option>
				<%
					sQuery = "Select Left(AccountDescription,45),AccountHead From VwAccheadforPurchaseApp  "&_
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
					Response.Write sAccHead &" " & objRs(1) & "<br>"
					IF CStr(sAccHead) = CStr(objRs(1)) Then
%>
						<option value="<%=objRs(1)%>" Selected><%=objRs(0)%></option>
<%					Else		%>
						<option value="<%=objRs(1)%>"><%=objRs(0)%></option>
<%
					End IF

					objRs.MoveNext
					Loop
					objRs.Close
%>
				</Select>
				</td>
				<td class="ExcelSerial" align="center" colspan="3"><%=sTaxName%>&nbsp;</td>
				<%if sTaxMode="F" then %>
				<td class="ExcelDisplayCell" align="right"></td>
				<%else%>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxPer<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxValue%>" onBlur="setTaxPercentage('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="5" size="7" class="Formelem">&nbsp;%</td>
				<%end if%>
				<td class="ExcelInputCell" align="right"><input type="text" style="text-align: Right" NAME="txtTaxValue<%=sCatCode%><%=sTaxCode%>" value="<%=dTaxAmt%>"  onBlur="setTaxAmount('<%=sCatCode%>','<%=sTaxCode%>',this)" Maxlength="13" size="15" class="Formelem"></td>
				    </tr>
			<%

	next
	 'Response.Write "iRoundedoff="&iRoundedoff    &"<BR><BR>"
 	IF CStr(sInvRndChk) = "1" Then
		iRoundedInvvalue = round(dInvamount)
		iRoundedoff = Round(cdbl(iRoundedInvvalue) - cdbl(dInvamount),2)
		'Response.Write "iRoundedoff"&iRoundedoff &"<BR><BR>"
	Else
		iRoundedInvvalue = dInvamount
		iRoundedoff = 0
	End IF

	Dim OldTaxNode,sExp
	sExp = "//TaxDetails"
	Set OldTaxNode = oNodRoot.selectNodes(sExp)

	'Response.Write "OldTaxNode.length="&
	IF OldTaxNode.length <> 0 Then
		OldTaxNode.removeAll
	End IF

	'Response.Write "iRoundedoff="&iRoundedoff
	oNodTaxRoot.Attributes.GetNamedItem("NettValue").value=iRoundedInvvalue
	oNodTaxRoot.Attributes.GetNamedItem("RoundOffValue").value=iRoundedoff
	oNodRoot.appendChild oNodTaxRoot

	oNodTaxRoot.Attributes.GetNamedItem("InvoiceVlaue").value=iRoundedInvvalue
	oNodRoot.appendChild oNodTaxRoot

	oDOM.save server.MapPath("../temp/transaction/Voucher Amd_PUR_"&Session.SessionID&".xml")
	'Response.Write dRndOff
	
	If dRndOff <> "" then dRndOff = FormatNumber(dRndOff,2,,,0)
%>
<tr>

    <td class="ExcelSerial" align="center" colspan="7">
    <p align="right"><b>Rounded Off&nbsp; </b></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtroundoff"  size="11" value="<%=dRndOff%>"  class="Formelem" readonly>
        </tr>


        <tr>

    <td class="ExcelSerial" align="right" colspan="7"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelInputCell" align="right"> <input type="text" style="text-align: Right" NAME="txtInvValue"  Maxlength="13" size="15" value="<%=FormatNumber(iRoundedInvvalue,2,,,0)%>" class="Formelem">
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
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td>
												 <% IF CStr(sAppValue) = "Y" Then %>
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
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="Cancel('VouPURBookSelection.asp')" class="ActionButton" >
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

