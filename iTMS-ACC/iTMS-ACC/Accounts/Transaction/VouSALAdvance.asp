<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouSALAdvance.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February  18, 2003
	'Modified By				:	Manohar Prabhu.R
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
<%
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTemp
dim sSalType,sOrgId,sQuery,sPartyName,sRefernceNo,sorgName,sInvDet
dim sParCode,sParSubType,sParType,dBasicValue,dInvAmount,dQty
dim bCommFlag,dNettAmount,sAdjCheck,sCrVouDate
Dim oNodItem
dim iItemCount,dTemp,dItemTotal,dItemValue,oNodTax
dim saTemp
dim dBasItemValue,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax

sAdjCheck = "F" ' Assign that no adjustemnt is not there

dQty=0
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")

oDOM.load server.MapPath("../temp/transaction/Voucher Entry_SAL_"&Session.SessionID&".xml")
set oNodRoot=oDOM.documentElement


bCommFlag=false
for each oNodDeatils in oNodRoot.childNodes
	if oNodDeatils.nodeName="Header" then
		for Each oNodEntry in  oNodDeatils.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.GetNamedItem("OrgId").value
				sorgName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SalesType" then
				sSalType=oNodEntry.Attributes.GetNamedItem("SalType").value
			end if
			if oNodEntry.nodeName="Party" then
				sParType=oNodEntry.Attributes.GetNamedItem("ParType").value
				sParSubType=oNodEntry.Attributes.GetNamedItem("ParSubType").value
				sParCode=oNodEntry.Attributes.GetNamedItem("ParCode").value
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sRefernceNo = oNodEntry.Attributes.GetNamedItem("RefNo").value
				sInvDet = oNodEntry.Attributes.GetNamedItem("InvNo").value &" - "&oNodEntry.Attributes.GetNamedItem("InvDate").value
				sVouDate = oNodEntry.Attributes.GetNamedItem("InvDate").value
				sCrVouDate = Trim(oNodEntry.Attributes.GetNamedItem("InvDate").value)
			end if
		next
	end if
	if oNodDeatils.nodeName="Details" then
		dBasicValue=oNodDeatils.Attributes.GetNamedItem("ActualValue").value
		for Each oNodEntry in  oNodDeatils.childNodes
			if oNodEntry.nodeName="Entry" then
				dQty=CDbl(dQty)+CDbl(oNodEntry.Attributes.GetNamedItem("Qty").value)
			end if
		next
		set oNodItem=oNodDeatils
	end if
	if oNodDeatils.nodeName="TaxDetails" then
		dInvAmount=oNodDeatils.Attributes.Item(0).nodeValue
		dNettAmount=oNodDeatils.Attributes.getNamedItem("NettValue").value
		Set oNodTax = oNodDeatils
	end if
	if oNodDeatils.nodeName="AgentDetails" then
		set oNodAgent=oNodDeatils
		bCommFlag=true
	end if

next



'------TO UPDATE THE ITEM VALUE FOR THE TAX WITHOUT ACCOUNT HEAD------


dTemp=0
dItemTotal = 0
dBasItemValue = 0

Dim sExp,iCtr,TempNode

sExp = "//Tax[@AccHead=0]"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		dTax = TempNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value
		dTaxValue = CDbl(dTaxValue) + CDbl(dTax)
	Next
End IF

sExp = "//Entry"
Set TempNode = oNodRoot.selectNodes(sExp)
iItemCount = TempNode.length

Dim dEachTaxVal,dNewTaxval,dDiffVal,dAddVal
dEachTaxVal = Round(CDbl(dTaxValue)/iItemCount,0)
dNewTaxval = Cdbl(dEachTaxVal) * CDbl(iItemCount)
IF CDbl(dNewTaxval) <> CDbl(dTaxValue) Then
	dDiffVal = Cdbl(dTaxValue) - CDbl(dNewTaxval)
Else
	dDiffVal = 0
End IF

'Response.Write dEachTaxVal &" " & dDiffVal &"<br>"

For iCtr = 0 To TempNode.length - 1
		dAddVal = 0
	IF CInt(iCtr) = 0 Then
		dAddVal = CDbl(dEachTaxVal) + CDbl(dDiffVal)
		dAddVal = CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value) + dAddVal
	Else
		dAddVal = CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value) + dEachTaxVal
	End IF
	'Response.Write dAddVal &"<br>"
	TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value = dAddVal
Next

oDOM.save server.MapPath("../temp/transaction/Voucher Entry_Sal_"&Session.SessionID&".xml")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<XML id="AdvanceData" src="<%="../temp/transaction/Voucher Entry_SAL_"&Session.SessionID&".xml"%>"></XML>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/Cancel.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/AdvanceAdjustmentCompat.js"></SCRIPT>
<script>
ITMSAdvanceAdjustmentCompat.install({
	invoiceAmount: "<%=dNettAmount%>",
	saveUrl: "XMLSave.asp?Mod=SAL&Name=Voucher Entry",
	disableButton: "btnNext",
	subtractToAccount: true,
	confirmNoAdjustment: true,
	confirmNoAdjustmentMessage: "Continue! Without Adjusting Advances?"
});
</script>
</HEAD>
<!--BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetAmount()" -->
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="VouSALGenerate.asp">
<Input type="hidden" name="hInvVal" value="<%=dNettAmount%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Voucher  	</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
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
								<td class="TabCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<td align="center">Invoice Details</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="100">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Commission</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="75">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
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
                            <td align="center" colspan="3" class="MiddlePack" height="7">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                            <tr>
                            <td align="center" height="39">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center" height="39">
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
                            <td align="center" height="39">
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                                            <table border="0" cellspacing="0"  cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub">Unit</td>
                                                    <td  class="FieldCellSub"  colspan="2"><span class="DataOnly"><%=sorgName%>&nbsp;</span> </td>
                                                        </tr>

                                                        <tr>
                                                    <td class="FieldCellSub">Invoice No-Date</td>
                                                    <td width="145" class="FieldCellSub"><span class="DataOnly"><%=sInvDet%>&nbsp;</span></td>
                                                    <td class="FieldCellSub">Reference No</td>
                                                    <td class="FieldCellSub" width="145"><span class="DataOnly"><%=sRefernceNo%>&nbsp;</span></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub">Invoice Total</td>
                                                    <td width="145" class="FieldCellSub"><span class="DataOnly"><%=FormatNumber(dNettAmount,2,,,0)%>&nbsp;</span></td>
                                                    <td class="FieldCellSub">Invoice Value</td>
                                                    <td class="FieldCellSub" width="145"><span class="DataOnly"><%=FormatNumber(dInvAmount,2,,,0)%>&nbsp;</span></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub">Party Name</td>
                                                    <td  class="FieldCellSub" colspan="2"><span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
                                                        </tr>

                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell">
												    <div align="left">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="125"><p align="center">Advance
                                                              Details
                                                            </td>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack> </td>
														</tr>

														<tr>
															<td class=FieldCellSub width="10">
												<DIV class=frmBody id=frm3 style="width: 580; height:220;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="600">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center"  rowspan="2">&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center" colspan="3">Document</td>
                                        <td class="ExcelHeaderCell" align="center" colspan="4">Amount</td>

                                            </tr>
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center">Number</td>
                                        <td class="ExcelHeaderCell" align="center">Date</td>
                                        <!--td class="ExcelHeaderCell" align="center">Detail</td-->
                                        <td class="ExcelHeaderCell" align="center">Amount</td>
                                        <td class="ExcelHeaderCell" align="center">To Account</td>
                                        <td class="ExcelHeaderCell" align="center">Adjusted</td>
                                        <td class="ExcelHeaderCell" align="center">To Adjust</td>
                                        <td class="ExcelHeaderCell" align="center">To be adjusted</td>
                                            </tr>
<%
dim iDocNo,iSno,oNodAdvRoot,newElem
dim dAmtReceived,dAmtAdjusted,sVouDate,sVouNo,iCrDocNo,dToAccount,dToAdjust
dim bAdvFlag,dAdvNo
dim sVouType,sRecdFrom,sInstType,sInstNo,sInstDate,sInstBank,sInstDet
Set oNodAdvRoot = oDOM.createElement("AdvanceDetails")
bAdvFlag=true
sAdjCheck = "F"

'sQuery="select a.TransactionNumber,b.VoucherNumber,convert(char,b.VoucherDate,103),"&_
'	"isnull(a.AdvanceAdjusted,0),isnull(a.AdvanceReceived,0),b.TransactionType,b.PayToRecdFrom,"&_
'	"b.BankInstrumentType,b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, "&_
'	"A.CreatedTransNo,A.AdvanceNumber  from Acc_T_AdvancePayments a ,Acc_T_VoucherHeader b" &_
'	" where a.OUDefinitionID='"&sOrgId&"' and a.PartyType='"&sParType&"'"&_
'	" and a.PartySubType="&sParSubType&" and a.PartyCode="&sParCode&" and "&_
'	"isnull(a.AdvanceReceived,0)>isnull(a.AdvanceAdjusted,0)  and b.TransactionNumber=a.TransactionNumber"&_
'	" and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103)"

'==========================================================================================
'Changed on		:	10/12/2004
'Changed By		:	Manohar Prabhu.R
'Reason			:	To also get the Created advance value also.
'==========================================================================================

sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(char,b.VoucherDate,103),isnull(a.AdvanceAdjusted,0), "&_
		 "isnull(a.AdvanceReceived,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
		 "b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
		 "A.AdvanceNumber,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
		 "from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
		 "a.OUDefinitionID='"&sOrgId&"' and a.PartyType='"&sParType&"' and a.PartySubType="&sParSubType&" and a.PartyCode="&sParCode&" and "&_
		 "isnull(a.AdvanceReceived,0)>isnull(a.AdvanceAdjusted,0) and "&_
		 "b.TransactionNumber = a.TransactionNumber "&_
		 "and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
		 "A.CreatedAdvanceNo = C.CreatedAdvanceNo "

'Response.Write sQuery &"<br>"

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

set iDocNo = objRs(0)
set sVouNo = objRs(1)
set sVouDate = objRs(2)
set sVouType = objRs(5)
set sRecdFrom = objRs(6)
set sInstType = objRs(7)
set sInstNo = objRs(8)
set sInstDate = objRs(9)
set sInstBank = objRs(10)
Set iCrDocNo = objRs(11)
Set dAdvNo = objRs(12)


iSno=1


Do While Not objRs.EOF
	sAdjCheck = "T"

	dAmtAdjusted = Trim(objRs(3))
	dAmtReceived = Trim(objRs(4))
	dToAccount = Trim(objRs(13))

	dAmtAdjusted = CDbl(dAmtAdjusted)
	dAmtReceived = CDbl(dAmtReceived)
	dToAccount = CDbl(dToAccount)

	dToAdjust = CDbl(dAmtReceived - dAmtAdjusted - dToAccount)

	IF CDbl(dToAdjust) > 0 Then
		Set newElem = oDOM.createElement("Advance")
		newElem.setAttribute "TransNo",iDocNo
		newElem.setAttribute "VoucherNo",trim(sVouNo)
		newElem.setAttribute "VoucherDate",trim(sVouDate)
		newElem.setAttribute "AmountRec",dAmtReceived
		newElem.setAttribute "AmountAdj",dAmtAdjusted
		newElem.setAttribute "AmountToAdj","0"
		newElem.setAttribute "CreatedTransNo",iCrDocNo
		newElem.setAttribute "AdvNo",dAdvNo
		newElem.setAttribute "AdjType","B"
		newElem.setAttribute "ToAccount", dToAccount


		oNodAdvRoot.appendChild newElem
		if sVouType="CAR" then
			sInstDet="Received From : &nbsp; " &sRecdFrom
		else
			sInstDet=sInstType
			sInstDet=sInstDet &"&nbsp; No:&nbsp;"& sInstNo
			sInstDet=sInstDet &"<br> Date:&nbsp;"& sInstDate
			sInstDet=sInstDet &"<br>Bank:&nbsp;"& sInstBank
		end if

%>
		              <tr>
			<td class="ExcelSerial" align="center"><%=iSno%></td>
			<td class="ExcelInputCell" align="right" width="10">
			<input type="checkbox" name="chkDocument<%=iDocNo%>Z<%=dAdvNo%>" value="<%=iDocNo%>" class="FormElem"></td>
			<td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
			<td class="ExcelDisplayCell"><p align="center"><%=sVouDate%></td>
			<!--td class="ExcelDisplayCell" align="left"><%=sInstDet%></td-->
			<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtReceived,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjust,2,,,0)%></td>

			<td class="ExcelInputCell" align="right">
			<input type="text" style="text-align: Right"  name="txtAmount<%=iDocNo%>Z<%=dAdvNo%>" value="0.00" size="15" maxlength="13" class="Formelem"> </td>
            </tr>

<%
		End IF
		objRs.MoveNext
		iSno=cint(iSno)+1
		Loop
	objRs.Close



	'====================================================================================================================================================
	'Added On		:	27/11/2004
	'Reason			:	To display all credit notes for the party.
	'====================================================================================================================================================
	'To display all sales return credit notes for the selected party,type and sub type.


	sQuery = "Select P.PayablesNumber,H.VoucherNumber,Convert(Char,P.PartyBillDate,103), "
	sQuery = sQuery & "Round(P.AmountPayable,0),P.AmountPaid,C.AmountPaid - P.AmountPaid ToAccount,P.PayablesNumber "
	sQuery = sQuery & "From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "
	sQuery = sQuery & "P.CreatedPayablesNumber = 0 and P.Narration is Null and P.AmountPayable > P.AmountPaid "
	sQuery = sQuery & "and P.OUDefinitionID = '"&sOrgId&"' and P.PartyType = 'DR' and P.PartyCode = "&sParCode&" and "
	sQuery = sQuery & "P.VoucherDate <= convert(datetime,'"&sCrVouDate&"',103) and P.TransactionNumber = "
	sQuery = sQuery & "H.TransactionNumber and C.PayablesNumber = P.CrCreatedPayable and Round(C.AmountPayable,2) > Round(C.AmountPaid,2)  "
	sQuery = sQuery & "and isNull(BankInstrumentType,'') = 'SR' "

	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End With
	Set objRs.ActiveConnection = Nothing

	Do While Not objRs.EOF

		sAdjCheck = "T"
		iDocNo = objRs(0)
		sVouNo = objRs(1)
		sVouDate = objRs(2)
		dAmtReceived = Trim(objRs(3))
		dAmtAdjusted = Trim(objRs(4))
		dToAccount = Trim(objRs(5))
		iCrDocNo = objRs(6)
		sInstDet = "Credit Note For Sales Return"
		sAdjCheck = "T"

		dToAdjust = 0
		dAmtReceived = CDbl(dAmtReceived)
		dAmtAdjusted = CDbl(dAmtAdjusted)
		dToAccount = CDbl(dToAccount)

		dToAdjust = CDbl(dAmtReceived - dAmtAdjusted - dToAccount)

		Set newElem = oDOM.createElement("Advance")
		newElem.setAttribute "TransNo",iDocNo
		newElem.setAttribute "VoucherNo",trim(sVouNo)
		newElem.setAttribute "VoucherDate",trim(sVouDate)
		newElem.setAttribute "AmountRec",dAmtReceived
		newElem.setAttribute "AmountAdj",dAmtAdjusted
		newElem.setAttribute "AmountToAdj","0"
		newElem.setAttribute "CreatedTransNo",iCrDocNo
		newElem.setAttribute "AdvNo","0"
		newElem.setAttribute "AdjType","SR"
		newElem.setAttribute "ToAccount", dToAccount
		oNodAdvRoot.appendChild newElem

%>
		<tr>
		<td class="ExcelSerial" align="center"><%=iSno%></td>
        <td class="ExcelInputCell" align="right" width="10">
        <input type="checkbox" name="chkDocument<%=iDocNo%>Z0" value="<%=iDocNo%>" class="FormElem"></td>
        <td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
        <td class="ExcelDisplayCell"><p align="center"><%=sVouDate%></td>
        <!--td class="ExcelDisplayCell" align="left"><%=sInstDet%></td-->
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtReceived,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjust,2,,,0)%></td>
        <td class="ExcelInputCell" align="right">
        <input type="text" style="text-align: Right"  name="txtAmount<%=iDocNo%>Z0" value="0.00" size="15" maxlength="13" class="Formelem"> </td>
            </tr>
 <%
		iSno = iSno + 1
		objRs.MoveNext
	loop
	objRs.Close



	'To display the list of Sales Commission for the selected party if any
	sQuery = "Select P.TransactionNumber,H.VoucherNumber,Convert(Char,P.PartyBillDate,103), "&_
			 "P.AmountPayable,P.AmountPaid,C.AmountPaid - P.AmountPaid ToAccount,P.PayablesNumber "&_
			 "From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "&_
			 "P.CreatedPayablesNumber = 0 and P.Narration is Null and P.AmountPayable > P.AmountPaid "&_
			 "and P.OUDefinitionID = '"&sOrgId&"' and P.PartyCode = "&sParCode&" and P.VoucherDate <= "&_
			 "convert(datetime,'"&sCrVouDate&"',103) and P.TransactionNumber = H.TransactionNumber "&_
			 "and H.BankInstrumentType = 'SC' and P.CrCreatedPayable = C.PayablesNumber "


	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End With
	Set objRs.ActiveConnection = Nothing

	'If not objRs.EOF then

	Do While Not objRs.EOF
		sAdjCheck = "T"
		iDocNo = objRs(0)
		sVouNo = objRs(1)
		sVouDate = objRs(2)
		dAmtReceived = Trim(objRs(3))
		dAmtAdjusted = Trim(objRs(4))
		dToAccount = Trim(objRs(5))
		iCrDocNo = objRs(6)
		sInstDet = "Sales Commission"
		sAdjCheck = "T"

		dAmtReceived = CDbl(dAmtReceived)
		dAmtAdjusted = CDbl(dAmtAdjusted)
		dToAccount = CDbl(dToAccount)

		dToAdjust = CDbl(dAmtReceived - dAmtAdjusted - dToAccount)

		Set newElem = oDOM.createElement("Advance")
		newElem.setAttribute "TransNo",iDocNo
		newElem.setAttribute "VoucherNo",trim(sVouNo)
		newElem.setAttribute "VoucherDate",trim(sVouDate)
		newElem.setAttribute "AmountRec",dAmtReceived
		newElem.setAttribute "AmountAdj",dAmtAdjusted
		newElem.setAttribute "AmountToAdj","0"
		newElem.setAttribute "CreatedTransNo",iCrDocNo
		newElem.setAttribute "AdvNo","0"
		newElem.setAttribute "AdjType","SC"
		newElem.setAttribute "ToAccount", dToAccount
		oNodAdvRoot.appendChild newElem

		sInstDet = "Sales Commission "

%>
		<tr>
		<td class="ExcelSerial" align="center"><%=iSno%></td>
        <td class="ExcelInputCell" align="right" width="10">
        <input type="checkbox" name="chkDocument<%=iDocNo%>Z0" value="<%=iDocNo%>" class="FormElem"></td>
        <td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
        <td class="ExcelDisplayCell"><p align="center"><%=sVouDate%></td>
        <!--td class="ExcelDisplayCell" align="left"><%=sInstDet%></td-->
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtReceived,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjust,2,,,0)%></td>
        <td class="ExcelInputCell" align="right">
        <input type="text" style="text-align: Right"  name="txtAmount<%=iDocNo%>Z0" value="0.00" size="15" maxlength="13" class="Formelem"> </td>
            </tr>

 <%
	objRs.MoveNext
	loop
	objRs.Close

	'To display the list of Credit Note of Type Other for the selected party if any
	sQuery = "Select P.TransactionNumber,H.VoucherNumber,Convert(Char,H.VoucherDate,103), "&_
			 "P.AmountPayable,P.AmountPaid,C.AmountPaid - P.AmountPaid ToAccount,P.PayablesNumber "&_
			 "From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "&_
			 "P.CreatedPayablesNumber = 0 and P.Narration is Null and P.AmountPayable > P.AmountPaid "&_
			 "and P.OUDefinitionID = '"&sOrgId&"' and P.PartyCode = "&sParCode&" and "&_
			 "P.VoucherDate <= convert(datetime,'"&sCrVouDate&"',103) and "&_
			 "P.TransactionNumber = H.TransactionNumber and H.BankInstrumentType = 'OT' "&_
			 "and P.CrCreatedPayable = C.PayablesNumber "




	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End With
	Set objRs.ActiveConnection = Nothing

	'If not objRs.EOF then

	Do While Not objRs.EOF
		sAdjCheck = "T"
		iDocNo = objRs(0)
		sVouNo = objRs(1)
		sVouDate = objRs(2)
		dAmtReceived = Trim(objRs(3))
		dAmtAdjusted = Trim(objRs(4))
		dToAccount = Trim(objRs(5))
		iCrDocNo = objRs(6)
		sInstDet = "Sales Commission"
		sAdjCheck = "T"

		dAmtReceived = CDbl(dAmtReceived)
		dAmtAdjusted = CDbl(dAmtAdjusted)
		dToAccount = CDbl(dToAccount)

		dToAdjust = CDbl(dAmtReceived - dAmtAdjusted - dToAccount)

		Set newElem = oDOM.createElement("Advance")
		newElem.setAttribute "TransNo",iDocNo
		newElem.setAttribute "VoucherNo",trim(sVouNo)
		newElem.setAttribute "VoucherDate",trim(sVouDate)
		newElem.setAttribute "AmountRec",dAmtReceived
		newElem.setAttribute "AmountAdj",dAmtAdjusted
		newElem.setAttribute "AmountToAdj","0"
		newElem.setAttribute "CreatedTransNo",iCrDocNo
		newElem.setAttribute "AdvNo","0"
		newElem.setAttribute "AdjType","OT"
		newElem.setAttribute "ToAccount", dToAccount
		oNodAdvRoot.appendChild newElem

		sInstDet = "Credit Note Others "

%>
		<tr>
		<td class="ExcelSerial" align="center"><%=iSno%></td>
        <td class="ExcelInputCell" align="right" width="10">
        <input type="checkbox" name="chkDocument<%=iDocNo%>Z0" value="<%=iDocNo%>" class="FormElem"></td>
        <td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
        <td class="ExcelDisplayCell"><p align="center"><%=sVouDate%></td>
        <!--td class="ExcelDisplayCell" align="left"><%=sInstDet%></td-->
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtReceived,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjust,2,,,0)%></td>
        <td class="ExcelInputCell" align="right">
        <input type="text" style="text-align: Right"  name="txtAmount<%=iDocNo%>Z0" value="0.00" size="15" maxlength="13" class="Formelem"> </td>
            </tr>

 <%
	objRs.MoveNext
	loop
	objRs.Close

%>

                                                </table>
												</div>
                                                   </td>
												</center>
														</tr>

														<tr>
															<td class=MiddlePack width="267">
                                                   </td>
														</tr>
													</table>
                                                            </td>
														</tr>
													</table>
                                                        </div>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
<%
Dim CheckNode
sExp = "//AdvanceDetails"
Set CheckNode = oNodRoot.selectNodes(sExp)

IF CStr(sAdjCheck) = "T" Then
	IF CheckNode.length <> 0 Then
		CheckNode.removeall()
	End IF
	oNodRoot.appendChild oNodAdvRoot
	oDOM.save server.MapPath("../temp/transaction/Voucher Entry_SAL_"&Session.SessionID&".xml")
else
	bAdvFlag=false
	Set oNodAdvRoot = oDOM.createElement("AdvanceDetails")
	oNodRoot.appendChild oNodAdvRoot
	oDOM.save server.MapPath("../temp/transaction/Voucher Entry_SAL_"&Session.SessionID&".xml")
	Response.Redirect("VouSALGenerate.asp")
end if


%>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell">
                                                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                            <tr>
                                        <td valign="middle" class="ActionCell"><p align="center">
                                        <input type="button" value="Next" name="btnNext" class="ActionButton" onClick="actionDone()">
                                        <input type="button" value="Cancel" name="B8" class="ActionButton" onClick="Cancel('../welcome_Accounts.asp')"></td>
                                            </tr>
                                                </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="BottomPack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
</HTML>