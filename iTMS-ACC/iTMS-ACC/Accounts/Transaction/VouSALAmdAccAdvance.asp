<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouSALAmdAccAdvance.asp
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
<%
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTemp
dim sSalType,sOrgId,sQuery,sPartyName,sRefernceNo,sorgName,sInvDet
dim sParCode,sParSubType,sParType,dBasicValue,dInvAmount,dQty
dim bCommFlag,dNettAmount,sCrVouDate,iCrTransNo,Objrs2

dQty=0
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")

oDOM.load server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")
set oNodRoot=oDOM.documentElement

'iCrTransNo = oNodRoot.Attributes.Item(0).nodeValue
iCrTransNo = Request("hTransNo")


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
				sRefernceNo=oNodEntry.Attributes.GetNamedItem("RefNo").value
				sInvDet=oNodEntry.Attributes.GetNamedItem("InvNo").value &" - "&oNodEntry.Attributes.GetNamedItem("InvDate").value
				sVouDate=oNodEntry.Attributes.GetNamedItem("InvDate").value
				sCrVouDate = oNodEntry.Attributes.GetNamedItem("InvDate").value
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

	end if
	if oNodDeatils.nodeName="TaxDetails" then
		dInvAmount=oNodDeatils.Attributes.Item(0).nodeValue
		dNettAmount=oNodDeatils.Attributes.getNamedItem("NettValue").value
	end if
	if oNodDeatils.nodeName="AgentDetails" then
		set oNodAgent=oNodDeatils
		bCommFlag=true
	end if

next


Dim sExp,iCtr,TempNode,dTax,dTaxValue,iItemCount

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

oDOM.save server.MapPath("../temp/transaction/Voucher AMD_Sal_"&Session.SessionID&".xml")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="AdvanceData" data-src="<%="../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"%>"></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/AdvanceAdjustmentCompat.js"></SCRIPT>
<script>
ITMSAdvanceAdjustmentCompat.install({
	invoiceAmount: "<%=dInvAmount%>",
	nettAmount: "<%=dNettAmount%>",
	saveUrl: "XMLSave.asp?Mod=SAL&Name=Voucher AMD",
	rowSuffix: true,
	includeCommission: true,
	availableMessage: "To be Adjusted Amount is Greater Than avilable Amount"
});
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="AmdAccSalGenerate.asp">
<Input type="hidden" name="hTransNo" Value="<%=iCrTransNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Accounted Sales Voucher Amendment 	</td>
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
                                                        <!--<tr>
                                                    <td class="FieldCellSub">Unit</td>
                                                    <td  class="FieldCellSub"  colspan="2"><span class="DataOnly"><%=sorgName%>&nbsp;</span> </td>
                                                        </tr>-->

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
<%if bCommFlag then%>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="125"><p align="center">Commission
                                                              Details
                                                            </td>
												</center>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable>
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack> </td>
														</tr>

														<tr>
															<td class=FieldCellSub width="10">
												<DIV class=frmBody id=frm3 style="width: 575; height:100;">
                                                                    <table border="0" id = "tblBin0" cellspacing ="1" class="ExcelTable" width="100%">
                                                                      <tr>
                                                                        <td class="ExcelHeaderCell" align="center" width="25">S.No.</td>
                                                                        <td class="ExcelHeaderCell" align="center">Agent Name</td>
                                                                        <td class="ExcelHeaderCell" align="center" colspan=3>Commission</td>
                                                                      </tr>
<%
dim sAgentName,sAgentCode,sCommType,dCommision,sComDetail,dCommisionValue
dim dInvValue
iSno=0
For Each oNodEntry in oNodAgent.childNodes
		iSno=CDbl(iSno)+1
		sAgentCode=oNodEntry.Attributes.Item(0).nodeValue
		sAgentName=oNodEntry.Attributes.Item(1).nodeValue
		sCommType=oNodEntry.Attributes.Item(2).nodeValue
		dCommision=oNodEntry.Attributes.Item(3).nodeValue

		if sCommType="Q" then
			sComDetail="On Qty -("&FormatNumber(dQty,3,,,0)&"* "&FormatNumber(dCommision,2,,,0)&")"
			dCommisionValue=CDbl(dCommision)*dQty
			dInvValue=dQty
		end if
		if sCommType="B" then
			sComDetail="On Basic Value -("&FormatNumber(dBasicValue,2,,,0)&"* "&FormatNumber(dCommision,2,,,0)& "%)"
			dCommisionValue=(CDbl(dCommision)*dBasicValue)/100
			dInvValue=dBasicValue
		end if
		if sCommType="V" then
			sComDetail="On Invoice Value("&FormatNumber(dInvAmount,2,,,0)&"* "&FormatNumber(dCommision,2,,,0)&" %)"
			dCommisionValue=(CDbl(dCommision)*dInvAmount)/100
			dInvValue=dInvAmount
		end if

		oNodEntry.Attributes.Item(4).nodeValue=dCommision
		dCommisionValue=FormatNumber(dCommisionValue,2,,,0)
%>
		<tr>
            <td class="ExcelSerial" align="center" width="25"><%=iSno%></td>
            <td class="ExcelDisplayCell" align="left"><%=sAgentName%></td>
            <td class="ExcelDisplayCell"  align="left"><%=sComDetail%>

            </td>

            <td class="ExcelDisplayCell" width="90" align="right">
            <input type="text" style="text-align:right" name="txtCommRate<%=sAgentCode%>" onBlur="setCommision(this,'<%=sAgentCode%>',<%=dInvValue%>)" value="<%=dCommision%>" size="14" class="Formelem">
            </td>
            <td class="ExcelInputCell" width="90" align="right">
            <input type="text" style="text-align:right" name="txtCommAmount<%=sAgentCode%>" value="<%=dCommisionValue%>" size="14" class="Formelem"></td>
          </tr>

<%next%>
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
<%end if%>
        <tr>
			<td align="center" colspan="3" class="MiddlePack">
                <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
			</td>
        </tr>

        <tr>
	<td align="center" width="5" class="ClearPixel"></td>
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
		<DIV class=frmBody id=frm3 style="width: 575; height:230;">
        <table border="0" cellspacing="1" class="ExcelTable" width="100%">
		    <tr>
		<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
		<td class="ExcelHeaderCell" align="center"  rowspan="2">&nbsp;</td>
		<td class="ExcelHeaderCell" align="center" colspan="4">Document</td>
		<td class="ExcelHeaderCell" align="center" colspan="3">Amount</td>

		    </tr>
            <tr>
        <td class="ExcelHeaderCell" align="center">Number</td>
        <td class="ExcelHeaderCell" align="center">Date</td>
        <td class="ExcelHeaderCell" align="center">Detail</td>
        <td class="ExcelHeaderCell" align="center">Amount</td>
        <td class="ExcelHeaderCell" align="center">Adjusted</td>
        <td class="ExcelHeaderCell" align="center">To Account</td>
        <td class="ExcelHeaderCell" align="center">To be adjusted</td>
            </tr>
<%
dim iDocNo,iSno,oNodAdvRoot,newElem,sAdjCheck
dim dAmtReceived,dAmtAdjusted,sVouDate,sVouNo
dim bAdvFlag,iAdjCrNo
dim sVouType,sRecdFrom,sInstType,sInstNo,sInstDate,sInstBank,sInstDet
Dim dToAccount,dTransAdjAmt,iCrAdvNo,iAdjTransNo
Set oNodAdvRoot = oDOM.createElement("AdvanceDetails")
bAdvFlag=true
iAdjCrNo = 0
sQuery = "SELECT A.TRANSACTIONNUMBER,B.VOUCHERNUMBER,CONVERT(CHAR,B.VOUCHERDATE,103), "&_
		 "ISNULL(A.ADVANCEADJUSTED,0),ISNULL(A.ADVANCERECEIVED,0),B.TRANSACTIONTYPE,B.PAYTORECDFROM, "&_
		 "B.BANKINSTRUMENTTYPE,B.BANKINSTRUMENTNO,CONVERT(CHAR,isNull(B.BANKINSTRUMENTDATE,''),103),DRAWNONBANK, "&_
		 "A.CREATEDTRANSNO, A.ADVANCENUMBER,ISNULL(C.ADVANCEADJUSTED,0) - ISNULL(A.ADVANCEADJUSTED,0) "&_
		 "TOACCOUNT,C.CREATEDADVANCENO  FROM ACC_T_ADVANCEPAYMENTS A, ACC_T_VOUCHERHEADER B, "&_
		 "ACC_T_CREATEDADVANCES C,ACC_T_CreatedAdvanceAdj J WHERE "&_
		 "A.OUDEFINITIONID='"&sOrgId&"' AND A.PARTYTYPE='"&sParType&"' AND "&_
		 "A.PARTYSUBTYPE="&sParSubType&" AND A.PARTYCODE="&sParCode&" AND B.TRANSACTIONNUMBER = A.TRANSACTIONNUMBER "&_
		 "AND B.VOUCHERDATE<= CONVERT(DATETIME,'"&sVouDate&"',103) AND "&_
		 "A.CREATEDADVANCENO = C.CREATEDADVANCENO AND C.CreatedAdvanceNo = J.CreatedAdvanceNo "&_
		 "and J.CreatedTransNo = "&iCrTransNo&" "

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing


iSno=1



Do While Not objRs.EOF
	iDocNo = objRs(0)
	sVouNo = objRs(1)
	sVouDate = objRs(2)
	dAmtAdjusted = objRs(3)
	dAmtReceived = objRs(4)
	sVouType = objRs(5)
	sRecdFrom = objRs(6)
	sInstType = objRs(7)
	sInstNo = objRs(8)
	sInstDate = objRs(9)
	sInstBank = objRs(10)
	dToAccount = objRs(13)
	iCrAdvNo = objRs(14)
	iAdjCrNo = iAdjCrNo &"," & iCrAdvNo

	IF Cstr(Trim(sInstDate)) = "01/01/1900" Then
		sInstDate = ""
	End IF


	sQuery = "Select AmountAdjusted From ACC_T_CreatedAdvanceAdj Where  "&_
			 "CreatedAdvanceNo = "&iCrAdvNo&" and CreatedTransNo = "&iCrTransNo&" "
	Objrs2.Open sQuery,Con
	IF Not Objrs2.EOF Then
		dTransAdjAmt = Objrs2(0)
	Else
		dTransAdjAmt = 0
	End IF
	Objrs2.Close

	dAmtAdjusted = CDbl(dAmtAdjusted) - Cdbl(dTransAdjAmt)
	dToAccount = CDbl(dToAccount) - CDbl(dTransAdjAmt)
	IF CDbl(dAmtReceived) > CDbl(dToAccount) Then
		Set newElem = oDOM.createElement("Advance")
		newElem.setAttribute "TransNo",iDocNo
		newElem.setAttribute "VoucherNo",trim(sVouNo)
		newElem.setAttribute "VoucherDate",trim(sVouDate)
		newElem.setAttribute "AmountRec",dAmtReceived
		newElem.setAttribute "AmountAdj",dAmtAdjusted
		newElem.setAttribute "AmountToAdj","0"
		newElem.setAttribute "ToAccount",dToAccount
		newElem.setAttribute "AdvNo",iCrAdvNo
		newElem.setAttribute "AdjType","B"

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
					<%IF CDbl(dTransAdjAmt) <> 0 Then %>
						<input type="checkbox" name="chkDocument<%=iDocNo%>Z<%=iCrAdvNo%>Z<%=iSno%>" value="<%=iDocNo%>" class="FormElem" Checked></td>
					<%Else%>
						<input type="checkbox" name="chkDocument<%=iDocNo%>Z<%=iCrAdvNo%>Z<%=iSno%>" value="<%=iDocNo%>" class="FormElem"></td>
					<%End IF %>
					<td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
					<td class="ExcelDisplayCell"><p align="center"><%=sVouDate%></td>
					<td class="ExcelDisplayCell" align="left"><%=sInstDet%></td>
					<td class="ExcelDisplayCell" align="right" ><%=FormatNumber(dAmtReceived,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
					<td class="ExcelInputCell" align="right">
					<input type="text" style="text-align: Right"  name="txtAmount<%=iDocNo%>Z<%=iCrAdvNo%>Z<%=iSno%>" value="<%=FormatNumber(dTransAdjAmt,2,,,0)%>" size="15" maxlength="13" class="Formelem"> </td>
				</tr>


<%
		End IF
		objRs.MoveNext
		iSno=cint(iSno)+1
	Loop
	objRs.Close

	sQuery = "SELECT A.TRANSACTIONNUMBER,B.VOUCHERNUMBER,CONVERT(CHAR,B.VOUCHERDATE,103), "
	sQuery = sQuery & "ISNULL(A.ADVANCEADJUSTED,0),ISNULL(A.ADVANCERECEIVED,0),B.TRANSACTIONTYPE,B.PAYTORECDFROM, "
	sQuery = sQuery & "B.BANKINSTRUMENTTYPE,B.BANKINSTRUMENTNO,CONVERT(CHAR,B.BANKINSTRUMENTDATE,103),DRAWNONBANK,  "
	sQuery = sQuery & "A.CREATEDTRANSNO, A.ADVANCENUMBER,ISNULL(C.ADVANCEADJUSTED,0) - ISNULL(A.ADVANCEADJUSTED,0)  "
	sQuery = sQuery & "TOACCOUNT,C.CREATEDADVANCENO  FROM ACC_T_ADVANCEPAYMENTS A, ACC_T_VOUCHERHEADER B,  "
	sQuery = sQuery & "ACC_T_CREATEDADVANCES C WHERE A.OUDEFINITIONID='"&sOrgId&"' AND A.PARTYTYPE='"&sParType&"' AND  "
	sQuery = sQuery & "A.PARTYSUBTYPE="&sParSubType&" AND A.PARTYCODE="&sParCode&" AND ISNULL(A.ADVANCERECEIVED,0)>ISNULL(A.ADVANCEADJUSTED,0)  "
	sQuery = sQuery & "AND B.TRANSACTIONNUMBER = A.TRANSACTIONNUMBER AND B.VOUCHERDATE <= CONVERT(DATETIME,'"&sVouDate&"',103) "
	sQuery = sQuery & "AND A.CREATEDADVANCENO = C.CREATEDADVANCENO   "
	sQuery = sQuery & "And ISNULL(A.ADVANCERECEIVED,0)>ISNULL(A.ADVANCEADJUSTED,0) and C.CREATEDADVANCENO NOT IN ("&iAdjCrNo&") "

	'Response.End

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
%>

<%




		Do While Not objRs.EOF
			iDocNo = objRs(0)
			sVouNo = objRs(1)
			sVouDate = objRs(2)
			dAmtAdjusted = objRs(3)
			dAmtReceived = objRs(4)
			sVouType = objRs(5)
			sRecdFrom = objRs(6)
			sInstType = objRs(7)
			sInstNo = objRs(8)
			sInstDate = objRs(9)
			sInstBank = objRs(10)
			dToAccount = objRs(13)
			iCrAdvNo = objRs(14)
			IF Cstr(Trim(sInstDate)) = "01/01/1900" Then
				sInstDate = ""
			End IF

			'IF CStr(Trim(iAdjTransNo)) = CStr(Trim(iCrTransNo)) or CStr(Trim(iAdjTransNo)) = "0" Then
				sQuery = "Select AmountAdjusted From ACC_T_CreatedAdvanceAdj Where  "&_
						 "CreatedAdvanceNo = "&iCrAdvNo&" and CreatedTransNo = "&iCrTransNo&" "
				Objrs2.Open sQuery,Con
				IF Not Objrs2.EOF Then
					dTransAdjAmt = Objrs2(0)
				Else
					dTransAdjAmt = 0
				End IF
				Objrs2.Close

				'Response.Write dToAccount &" -- " & dTransAdjAmt &"<br><br>"
				'IF CDbl(dToAccount) > 0 Then
					dToAccount = CDbl(dToAccount) - CDbl(dTransAdjAmt)
				'End IF
				'dToAccount = FormatNumber(dToAccount,2,,,0)

				'IF CDbl(dToAccount) > 0 Then

				IF CDbl(dAmtReceived) > CDbl(dToAccount) Then

					Set newElem = oDOM.createElement("Advance")
					newElem.setAttribute "TransNo",iDocNo
					newElem.setAttribute "VoucherNo",trim(sVouNo)
					newElem.setAttribute "VoucherDate",trim(sVouDate)
					newElem.setAttribute "AmountRec",dAmtReceived
					newElem.setAttribute "AmountAdj",dAmtAdjusted
					newElem.setAttribute "AmountToAdj","0"
					newElem.setAttribute "ToAccount",dToAccount
					newElem.setAttribute "AdvNo",iCrAdvNo
					newElem.setAttribute "AdjType","B"

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
					<%IF CDbl(dTransAdjAmt) <> 0 Then %>
						<input type="checkbox" name="chkDocument<%=iDocNo%>Z<%=iCrAdvNo%>Z<%=iSno%>" value="<%=iDocNo%>" class="FormElem" Checked></td>
					<%Else%>
						<input type="checkbox" name="chkDocument<%=iDocNo%>Z<%=iCrAdvNo%>Z<%=iSno%>" value="<%=iDocNo%>" class="FormElem"></td>
					<%End IF %>
						<td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
						<td class="ExcelDisplayCell"><p align="center"><%=sVouDate%></td>
						<td class="ExcelDisplayCell" align="left"><%=sInstDet%></td>
						<td class="ExcelDisplayCell" align="right" ><%=FormatNumber(dAmtReceived,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
						<td class="ExcelInputCell" align="right">
						<input type="text" style="text-align: Right"  name="txtAmount<%=iDocNo%>Z<%=iCrAdvNo%>Z<%=iSno%>" value="<%=FormatNumber(dTransAdjAmt,2,,,0)%>" size="15" maxlength="13" class="Formelem"> </td>
					</tr>


<%
				End IF
		'End IF
		objRs.MoveNext
		iSno=cint(iSno)+1
	Loop

	objRs.Close

'====================================================================================================================================================
'		Added On		:	27/11/2004
'		Reason			:	To display all credit notes for the party.
'====================================================================================================================================================
	'To display all sales return credit notes for the selected party,type and sub type.
	sQuery = "Select P.PayablesNumber,H.VoucherNumber,Convert(Char,P.PartyBillDate,103), "
	sQuery = sQuery & "Round(P.AmountPayable,0),P.AmountPaid,C.AmountPaid,P.AmountPayable - P.AmountPaid,P.PayablesNumber "
	sQuery = sQuery & "From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "
	sQuery = sQuery & "P.CreatedPayablesNumber = 0 and P.Narration is Null and P.AmountPayable > P.AmountPaid "
	sQuery = sQuery & "and P.OUDefinitionID = '"&sOrgId&"' and P.PartyType = 'DR' and P.PartyCode = "&sParCode&" and "
	sQuery = sQuery & "P.VoucherDate <= convert(datetime,'"&sCrVouDate&"',103) and P.TransactionNumber = "
	sQuery = sQuery & "H.TransactionNumber and C.PayablesNumber = P.CrCreatedPayable and Round(C.AmountPayable,2) > Round(C.AmountPaid,2) "

	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End With
	Set objRs.ActiveConnection = Nothing
	Dim iCrDocNo
	Do While Not objRs.EOF
		sAdjCheck = "T"
		iDocNo = objRs(0)
		sVouNo = objRs(1)
		sVouDate = objRs(2)
		dAmtReceived = objRs(3)
		dAmtAdjusted = objRs(5)
		iCrDocNo = objRs(7)
		sInstDet = "Credit Note For Sales Return"
		sAdjCheck = "T"

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
		oNodAdvRoot.appendChild newElem

%>
		<tr>
			<td class="ExcelSerial" align="center"><%=iSno%></td>
			<td class="ExcelInputCell" align="right" width="10">
			<input type="checkbox" name="chkDocument<%=iDocNo%>Z0Z<%=iSno%>" value="<%=iDocNo%>" class="FormElem"></td>
			<td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
			<td class="ExcelDisplayCell"><p align="center"><%=sVouDate%></td>
			<td class="ExcelDisplayCell" align="left"><%=sInstDet%></td>
			<td class="ExcelDisplayCell" ><%=FormatNumber(dAmtReceived,2,,,0)%></td>
			<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
			<td class="ExcelInputCell" align="right">
			<input type="text" style="text-align: Right"  name="txtAmount<%=iDocNo%>Z0Z<%=iSno%>" value="<%=FormatNumber(cdbl(dAmtReceived)-cdbl(dAmtAdjusted),2,,,0)%>" size="15" maxlength="13" class="Formelem"> </td>
        </tr>
 <%
		iSno = iSno + 1
		objRs.MoveNext
	loop
	objRs.Close



	'To display the list of Sales Commission for the selected party if any
	sQuery = "Select P.TransactionNumber,H.VoucherNumber,Convert(Char,P.PartyBillDate,103), "
	sQuery = sQuery & "P.AmountPayable,P.AmountPaid,P.AmountPaid,P.AmountPayable - P.AmountPaid,P.PayablesNumber "
	sQuery = sQuery & "From Acc_T_Payables P, Acc_T_VoucherHeader H Where P.CreatedPayablesNumber = 0 "
	sQuery = sQuery & "and P.Narration is Null and P.AmountPayable > P.AmountPaid and P.OUDefinitionID = '"&sOrgId&"' "
	sQuery = sQuery & "and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"&sCrVouDate&"',103) "
	sQuery = sQuery & "and P.TransactionNumber = H.TransactionNumber and H.BankInstrumentType = 'SC' "

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
		dAmtReceived = objRs(3)
		dAmtAdjusted = objRs(5)
		iCrDocNo = objRs(7)
		sInstDet = "Sales Commission"
		sAdjCheck = "T"

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
		oNodAdvRoot.appendChild newElem

		sInstDet = "Sales Commission "

%>
		<tr>
		<td class="ExcelSerial" align="center"><%=iSno%></td>
        <td class="ExcelInputCell" align="right" width="10">
        <input type="checkbox" name="chkDocument<%=iDocNo%>Z0Z<%=iSno%>" value="<%=iDocNo%>" class="FormElem"></td>
        <td class="ExcelDisplayCell" align="center"><%=sVouNo%></td>
        <td class="ExcelDisplayCell"><p align="center"><%=sVouDate%></td>
        <td class="ExcelDisplayCell" align="left"><%=sInstDet%></td>
        <td class="ExcelDisplayCell" ><%=FormatNumber(dAmtReceived,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dAmtAdjusted,2,,,0)%></td>
        <td class="ExcelDisplayCell" align="right">0.00</td>
        <td class="ExcelInputCell" align="right">
        <input type="text" style="text-align: Right"  name="txtAmount<%=iDocNo%>Z0Z<%=iSno%>" value="<%=FormatNumber(cdbl(dAmtReceived)-cdbl(dAmtAdjusted),2,,,0)%>" size="15" maxlength="13" class="Formelem"> </td>
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
'else

	'Response.Write "============== "

'	bAdvFlag=false
'	if bCommFlag=false then
'		sExp = "//AdvanceDetails"
'		Set TempNode = oNodRoot.selectNodes(sExp)
'		IF TempNode.length <> 0 Then
'			TempNode.removeAll
'		End IF
'		Set oNodAdvRoot = oDOM.createElement("AdvanceDetails")
'		oNodRoot.appendChild oNodAdvRoot
'		oDOM.save server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")
'
'		'Response.Redirect("VouSALAmdGenerate.asp")
'	end if
'
'end if
'objRs.Close
Dim iChkCtr
sExp = "//AdvanceDetails"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	'TempNode.removeAll
	For iChkCtr = 0 To TempNode.length - 1
		oNodRoot.RemoveChild TempNode.Item(iChkCtr)
	Next
End IF

oNodRoot.appendChild oNodAdvRoot
oDOM.save server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")
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
                                        <input type="button" value="Next" name="B7" class="ActionButton" onClick="actionDone()">
                                        <input type="button" value="Cancel" name="btnCancel" onClick="Cancel('VouPURBookSelection.asp')" class="ActionButton" >
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
