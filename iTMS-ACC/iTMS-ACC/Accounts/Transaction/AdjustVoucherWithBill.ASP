<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AdjustVoucherWithBill.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU
	'Created On					:	January 03,2002
	'Modified By				:	Manohar Prabhu.R
	'Modified On				:
	'Modified By				:	UmaMaheswari S
	'Modified On				:	20 June 2011
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<%
dim sOrgId,sPartyValue,objRs,sQuery,sVouType,sParCode,sParSubType,sParType,sVouDate
Dim sQuery2,sQuery3,sQuery4,objRs2,sTemp,iCrTransNo,iTrNo,oDOM,dAmount,sUnit,sParName
Dim iPiTrNo,iDNTrNo,iCNTrNo,iITrNo,iAdvRecTrNo,iAdvPayTrNo,sAdjTy,iCtr,sSelParTy
Dim sExp,TempNode,Root,sDispClass,sTransTy,sAmdTy,iAdvNo,sParChk,dAdjAmt,sPrintAdjusted

Dim blnSalesInvoice,blnDebitNode,blnPurBills,blnCreditNode,blnAdvancePay,blnAdvanceRec
Dim nVoucherNo,sVoucherDate,nVoucherNoQ
Dim nAmtOutstanding,nInvNumber,nBillNo,sCheck

sParChk = Request("sType")
sCheck = Request("hType")
'Response.Write "<p><font color=red>sCheck="&sCheck

dAdjAmt = 0.00
IF CStr(sParChk) = "" Then
	sTemp = Split(Request("Val"),"~")
	iCrTransNo = sTemp(0)
	sAmdTy = Request("AmdTy")
	IF CStr(sAmdTy) = "A" Then
		iAdvNo = sTemp(1)
	Else
		iAdvNo = 0
	End IF
Else
	'iAdvNo = 0
	iCrTransNo = 0
	sOrgId = Request.Form("SelOrg")
	sParCode = Request.Form("hdPartyCode")
	sUnit = Request.Form("hUnitName")
	sParName = Request.Form("hParName")
	sParType = Request.Form("SelPartyType")
End IF

IF CStr(iAdvNo) = "" Then
	iAdvNo = 0
	dAdjAmt = "0.00"
End IF


if trim(sVouDate)="" then
	sVouDate=Day(date)&"/"&Month(date)&"/"&Year(date)
end if

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")
'Response.Write "<p><font color=red>iCrTransNo="&iCrTransNo
IF CStr(iCrTransNo) <> 0 Then
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.load server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")

	sQuery = "Select H.TransactionNumber,D.AccUnitPartyType,D.AccUnitPartySubType,D.AccUnitPartyCode,   " &_
			 "H.OUDefinitionID,Convert(Varchar,H.VoucherDate,103) VoucherDate,H.VoucherAmount, "&_
			 "H.TransactionType,isNull(D.AccUnitPartyType,'NA'),AccUnitPartyType,H.VoucherNumber From " & _
			 "Acc_T_VoucherHeader H,Acc_T_VoucherDetails D Where H.TransactionNumber = D.TransactionNumber " & _
			 "And H.CreatedTransNo = " & iCrTransNo
	'Response.write "<p><font color=red>"&squery
	objRs.Open sQuery,con
	IF Not objRs.EOF Then
		iTrNo = objRs("TransactionNumber")
		sParType = objRs("AccUnitPartyType")
		sParSubType = objRs("AccUnitPartySubType")
		sParCode = objRs("AccUnitPartyCode")
		sOrgId = objRs("OUDefinitionID")
		sVouDate = objRs("VoucherDate")
		dAmount = objRs("VoucherAmount")
		sTransTy = objRs("TransactionType")
		sSelParTy = objRs("AccUnitPartyType")
		nVoucherNoQ = objRs("VoucherNumber")
	End IF
	objRs.Close
	'If sSelParTy = "" Then sSelParTy = "CR"
	'Response.write "<p>sSelParTy="&sSelParTy
	
	Set Root = oDOM.documentElement

	sExp = "//PayRec/Doc[@AdjType=""PI""]"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iPiTrNo = iPiTrNo &"," & TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		Next
		iPiTrNo = Mid(iPiTrNo,2)
	Else
		iPiTrNo = 0
	End IF

	sExp = "//PayRec/Doc[@AdjType=""I""]"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iITrNo = iITrNo &"," & TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		Next
		iITrNo = Mid(iITrNo,2)
	Else
		iITrNo = 0
	End IF

	sExp = "//PayRec/Doc[@AdjType=""C""]"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iCNTrNo = iCNTrNo &"," & TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		Next
		iCNTrNo = Mid(iCNTrNo,2)
	Else
		iCNTrNo = 0
	End IF

	sExp = "//PayRec/Doc[@AdjType=""D""]"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iDNTrNo = iDNTrNo &"," & TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		Next
		iDNTrNo = Mid(iDNTrNo,2)
	Else
		iDNTrNo = 0
	End IF


	sExp = "//PayRec/Doc[@AdjType=""R""]"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iAdvRecTrNo = iAdvRecTrNo &"," & TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		Next
		iAdvRecTrNo = Mid(iAdvRecTrNo,2)
	Else
		iAdvRecTrNo = 0
	End IF

	sExp = "//PayRec/Doc[@AdjType=""P""]"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			iAdvPayTrNo = iAdvPayTrNo &"," & TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		Next
		iAdvPayTrNo = Mid(iAdvPayTrNo,2)
	Else
		iAdvPayTrNo = 0
	End IF
Else
	iPiTrNo = 0
	iITrNo = 0
	iCNTrNo = 0
	iDNTrNo = 0
	iAdvRecTrNo = 0
	iAdvPayTrNo = 0
	dAmount = 0
End IF

IF CStr(iCrTransNo) = "0" and CStr(iAdvNo) <> "0" Then
	sQuery = "Select Top 1 B.OUDefinitionID,B.PartyCode,P.PartyName,D.OrgUnitShortDescription,B.VoucherAmt  "&_
			 "From Acc_T_BillsClosingDetails As B Inner Join APP_M_PartyMaster As P  "&_
			 "On B.PartyCode = P.PartyCode Inner Join DCS_OrganizationUnitDefinitions As D "&_
			 "On B.OUDefinitionID = D.OUDefinitionID Where B.AdvCloseNo = "&iAdvNo

	objRs.Open sQuery,con
	IF Not objRs.EOF Then
		sOrgId = objRs("OUDefinitionID")
		sParCode = objRs("PartyCode")
		sUnit = objRs("OrgUnitShortDescription")
		sParName = objRs("PartyName")
		dAdjAmt = FormatNumber(objRs("VoucherAmt"),2,,,0)
	End IF
	objRs.Close
End IF
'Response.Write Root.xml
If sCheck = "A" Then
	oDOM.load server.MapPath("../temp/transaction/Adjust_Det_"&Session.SessionID&".xml")
	Set Root = oDOM.documentElement
End IF
'Response.Write "<p><font color=red>iTrNo="&iTrNo & "," & iCrTransNo
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML id="AccHeadData">
<root No="0">
	<PayRec/>
	<RecCount Val="0" />
</root>
</XML>
<%If sCheck = "" Then%>
	<XML id="AdjDet" src="<%="../xmldata/Voucher/"&iCrTransNo&".xml"%>"></XML>
<%Else%>
	<XML id="AdjDet" src="<%="../temp/transaction/Adjust_Det_"&Session.SessionID&".xml"%>"></XML>
<%End IF%>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script SRC="../../scripts/AdjustVoucherWithBillCompat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetTotalVal('<%=iAdvNo%>')">
<!--<form method="POST" name="frm1" action="AdjustVouWithBillUpdate.asp">-->
<form method="POST" name="frm1" action="">
<input type=hidden name="hVouType" value="<%=sVouType%>">
<input type=hidden name="hCrTrNo" value="<%=iCrTransNo%>">
<input type=hidden name="hTrNo" value="<%=iTrNo%>">
<input type=hidden name="hVouAmt" value="<%=dAmount%>">
<input type=hidden name="hTransTy" value="<%=sTransTy%>">
<input type=hidden name="hAdvNo" value="<%=iAdvNo%>">
<input type=hidden name="hParCode" value="<%=sParCode%>">
<input type=hidden name="hUnitID" value="<%=sOrgId%>">
<input type=hidden name="hVouDate" value="<%=sVouDate%>">
<input type=hidden name="hDrCrAmt" value="0">
<input type=hidden name="hType" value="<%=scheck%>">
<input type=hidden name="hVouNumber" value="<%=nVoucherNoQ%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		Adjustments
		</td>
    </tr>
     <%IF CStr(sSelParTy) <> "NA" Then %>
	<tr>
		<td align="center" class="TopPack">
			<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>

							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
								<!--
								<%'IF Cstr(iCrTransNo) = "0" Then %>
									<DIV class=frmBody id=frm1 width="100%" style="height:75;">
								<%'Else%>
									<DIV class=frmBody id=frm1 width="100%" style="height:175;">
								<%'End IF%>
								-->

								<DIV class=frmBody id=frm1 width="100%" >
												<table border="0" cellpadding="0" cellspacing="1" width="100%" class="TableOutlineOnly">
								<%
									'IF CStr(iCrTransNo) <> 0 Then
										sExp = "//voucher"
										Set TempNode = Root.selectNodes(sExp)
										nVoucherNo = TempNode.Item(0).Attributes.getNamedItem("VoucherNo").Value
										If nVoucherNo = "" Then nVoucherNo = cdbl(nVoucherNoQ)
										sVoucherDate = TempNode.Item(0).Attributes.getNamedItem("VouDate").Value
								%>
											<!--<tr>
											  <td width="25%" class="FieldCellSub">Unit Name</td>
											  <td width="25%" class="FieldCellSub"><span class="DataOnly"><%=TempNode.Item(0).Attributes.getNamedItem("UnitName").Value%> </span></td>
											  <td width="25%" class="FieldCellSub">Voucher Number</td>
											  <td width="25%" class="FieldCellSub"><span class="DataOnly"><%=TempNode.Item(0).Attributes.getNamedItem("VoucherNo").Value%> </span></td>
											</tr>
											<tr>
											  <td width="25%" class="FieldCellSub">Voucher Date</td>
											  <td width="25%" class="FieldCellSub"><span class="DataOnly"><%=TempNode.Item(0).Attributes.getNamedItem("VouDate").Value%> </span></td>
											  <td width="25%" class="FieldCellSub">Amount</td>
											  <td width="25%" class="FieldCellSub"><span class="DataOnly" id="spVouAmt"><%=FormatNumber(dAmount,2,,,0)%> </span></td>
											</tr>-->
										   <tr>
											  <td width="25%" class="FieldCellSub">Book Name</td>
											  <td width="45%" class="FieldCellSub"><span class="DataOnly"><%=TempNode.Item(0).Attributes.getNamedItem("BookName").Value%> </span></td>
											  <td width="25%" class="FieldCellSub">Voucher Number</td>
											  <td width="10%" class="FieldCellSub">
												<span class="DataOnly"><%=nVoucherNo%> </span> &nbsp;&nbsp;
												<img src="../../assets/images/iTMS%20icons/DetailsIcon.gif" onClick="ShowVouch(<%=iCrTransNo%>)" alt="View Voucher Details" style="cursor: pointer" >
											  </td>
											</tr>
											<%
												sExp = "//AccHead "
												Set TempNode = Root.selectNodes(sExp)
											%>

										  <tr>
										    <td class="FieldCellSub">Party Name</td>
										    <td width="45%" class="FieldCellSub"><span class="DataOnly"><%=TempNode.Item(0).Attributes.getNamedItem("Name").Value%> </span></td>
										    <td width="15%" class="FieldCellSub">Voucher Date</td>
											<td width="15%" class="FieldCellSub"><span class="DataOnly"><%=sVoucherDate%> </span></td>
										  </tr>
										  <tr>
										    <td class="FieldCellSub">Party Type</td>
										    <td width="45%" class="FieldCellSub"><span class="DataOnly">
										    <%
												sTemp = Split(TempNode.Item(0).Attributes.getNamedItem("No").Value,"?")
												Response.Write(sTemp(2))
											%>
										    </span></td>
										    <td width="15%" class="FieldCellSub">Amount</td>
											<td width="15%" class="FieldCellSub"><span class="DataOnly" id="spVouAmt"><%=FormatNumber(dAmount,2,,,0)%> </span></td>
										  </tr>

										  <!--<tr>
										    <td class="FieldCellSub">Amount Adjusted</td>
										    <td colspan="2" class="FieldCellSub">
										    <%IF CStr(iAdvNo) = "0" Then %>
												<span class="DataOnly" id="spAmtVal">0.00</span>
											<%Else%>
												<span class="DataOnly" id="spAmtVal"><%=FormatNumber(dAmount,2,,,0)%></span>
											<%End IF %>

										    </td>
										  </tr>-->

										  <tr>
										    <td width="100%" class="FieldCellSub" colspan="2">Bill Adjustment Details</td>
										  </tr>

										  <tr>
										    <td width="100%" class="FieldCellSub" colspan="4">
										      <table border="0" cellpadding="0" cellspacing="1" width="100%" class="ExcelTable">
										        <tr>
										          <td width="2%" class="ExcelHeaderCell" align="center">S.No</td>
										          <td class="ExcelHeaderCell" align="center">
													<img border="0" src="../../assets/images/iTMS Icons/DeleteIcon.gif" alt="Remove Bill Details" onclick="DeleteData()"></a>	
												  </td>
										          <td width="55%" class="ExcelHeaderCell" align="center">Bill Details</td>
										          <td width="19%" class="ExcelHeaderCell" align="center">Bill Amount</td>
										          <td width="16%" class="ExcelHeaderCell" align="center">Adjusted Amount</td>
										        </tr>
										        <%	Dim nAmtToAdjust,sCheckDelAtt,nTotAmtAdjusted
													nTotAmtAdjusted = cdbl("0")
													sExp = "//PayRec/Doc"
													Set TempNode = Root.selectNodes(sExp)
													For iCtr = 0 To TempNode.length - 1
														sCheckDelAtt = ""
														sAdjTy = TempNode.Item(iCtr).Attributes.getNamedItem("AdjType").Value
														'Response.Write "<p><font color=red>sAdjTy="&sAdjTy
														If InStr(1,TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value,":") Then
															nInvNumber = split(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value,":")(1)
															iPlace = InStr(1,nInvNumber,"Dt")
															nInvNumber = Trim(Mid(nInvNumber,1,iPlace-1))
														Else
															nInvNumber = TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value
														End IF
														
														nAmtToAdjust = FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value,2,,,0)
														If sCheck = "A" Then
															sCheckDelAtt = TempNode.Item(iCtr).Attributes.getNamedItem("Check").Value
														End If
														'Response.Write "<p>sCheckDelAtt="&sCheckDelAtt
														If sCheckDelAtt <> "DEL" Then 
														

							%>
											<tr>
												<td class="ExcelHeaderCell" align="center"><%Response.Write(CInt(iCtr)+1)%></td>
												<td class="ExcelHeaderCell" align="center">
													<Input type="Checkbox" Name="ChkBillDet"  value="<%=nInvNumber&":"&nAmtToAdjust%>" >
												</td>
												<%IF CStr(sAdjTy) = "I" Then %>
													<td class="ExcelDisplayCell" align="left">
													<%
														IF Cstr(Trim(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value)) = "" Then
															sQuery = "Select isNull(Narration,'') From Acc_T_Receivables Where CreatedReceivable = "&TempNode.Item(iCtr).Attributes.Item(0).nodeValue
															'Response.write sQuery
															Objrs.Open sQuery,Con
															IF Not Objrs.Eof Then
																Response.Write Objrs(0)
															End IF
															objrs.close
														Else
															Response.Write(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value)
														End IF
													%>

													</td>
												<%Elseif CStr(sAdjTy) = "PI" Then %>
													<td class="ExcelDisplayCell" align="left">
														<%Response.Write(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value)%>
													</td>
												<%Else%>
													<td class="ExcelDisplayCell" align="left">
														<%Response.Write(TempNode.Item(iCtr).Attributes.getNamedItem("InvNo").Value)%> Dt:
														<%Response.Write(TempNode.Item(iCtr).Attributes.getNamedItem("InvDate").Value)%>
													</td>
												<%End IF %>
												<td class="ExcelDisplayCell" align="right"><%Response.Write(FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("TransAmount").Value,2,,,0))%></td>
												<td class="ExcelDisplayCell" align="right"><%Response.Write(FormatNumber(TempNode.Item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value,2,,,0))%></td>
											</tr>
										<%	nTotAmtAdjusted = cdbl(nTotAmtAdjusted) + cdbl(TempNode.Item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value)
										End IF 'If sCheckDelAtt <> "DEL" Then 
										Next
										%>

										<%'Response.Write "<p>nTotAmtAdjusted="&nTotAmtAdjusted
										If sCheckDelAtt = "DEL" Then iCtr = 0
										IF CStr(iCtr) = "0" Then
										%>

											<tr>
												<td class="ExcelDisplayCell" align="center" colspan="5"><b>No Bills is Selected for Adjustments!! Adjustment with New Bills is Not Possible.</b></td>
											</tr>

										<%End IF %>
										<%'IF CStr(iCtr) <> "0" Then%>
											<tr>
												<td class="ExcelDisplayCell" colspan="4" align="Right"><b>Total Adjusted Amount</b></td>
												<td class="ExcelDisplayCell" align="right">
													<Input type="Text" Name="TxtTotAdjustedAmount" style="text-align:right" Value="<%=FormatNumber(nTotAmtAdjusted,2,,,0)%>" Readonly Class="FormElemRead">
												</td>
											</tr>
										<%'End IF %>
										
										      </table>
										    </td>
										  </tr>
										  
										</table>
									<table border="0" cellspacing="0" width="100%">
									 <tr>
										<td align="center" class="MiddlePack" colspan="3">
										</td>
									</tr>
								<%'Else%>
										<!--<tr>
										  <td width="25%" class="FieldCellSub">Unit Name</td>
										  <td width="25%" class="FieldCellSub"><span class="DataOnly"><%=sUnit%> </span></td>
										  <td width="25%" class="FieldCellSub">Party Type</td>
										  <td colspan="3" class="FieldCellSub"><span class="DataOnly"><span id="spParTy"><%=sParType%> </span></td>
 										</tr>
										<tr>
										  <td width="25%" class="FieldCellSub">Party Name</td>
										  <td colspan="3" class="FieldCellSub"><span class="DataOnly"><span><%=sParName%> </span></td>
										</tr>-->
										<tr>
										  <td width="25%" class="FieldCellSub">Debit Amount</td>
										  <td Width="25%" class="FieldCellSub"><span class="DataOnly" id="spDRAmt"><%=dAdjAmt%> </span></td>
										  <td width="25%" class="FieldCellSub">Credit Amount</td>
										  <td Width="25%" class="FieldCellSub"><span class="DataOnly" id="spCRAmt"><%=dAdjAmt%> </span></td>
										</tr>
										<tr>
										  <td width="25%" class="FieldCellSub">Differences</td>
										  <td Width="25%" class="FieldCellSub"><span class="DataOnly" id="spDiffAmt">0.00 </span></td>
										  <td width="25%" class="FieldCellSub">&nbsp;</td>
										  <td Width="25%" class="FieldCellSub">&nbsp;</td>
										</tr>
								<%'End IF %>
									</table>
									</Div>
									<DIV class=frmBody id=frm1 width="100%" style="height:330;">
								<%' IF CStr(iCtr) <> "0" Then %>
                                                <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                               <tr>
													<td class="ExcelHeaderCell" align="Left" colspan="81">Bills Type</td>
											   </tr>
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center"  rowspan="2"></td>
                                        <td class="ExcelHeaderCell" align="center" colspan="3">Document</td>
                                        <td class="ExcelHeaderCell" align="center" colspan="4">Amount </td>

                                            </tr>
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center">Number</td>
                                        <td class="ExcelHeaderCell" align="center">Date</td>
                                        <td class="ExcelHeaderCell" align="center">Amount</td>
                                        <td class="ExcelHeaderCell" align="center"> Adjusted</td>
                                        <% 'IF CStr(sVouType) = "D" Then %>
                                        <!--td class="ExcelHeaderCell" align="center"> Cr Note Value</td>
                                        <%'Else%>
                                        <td class="ExcelHeaderCell" align="center"> Dr Note Value</td-->
                                        <%'End IF %>

                                        <td class="ExcelHeaderCell" align="center">To Account</td>
                                        <td class="ExcelHeaderCell" align="center">Outstanding</td>
                                        <td class="ExcelHeaderCell" align="center">To Adjust</td>

                                            </tr>
 <tr><td colspan="9" class="ExcelSerial"><b>Sales Invoices</b></td></tr>

<%

'*********************** Adjustment Type Details ******************************************
'		I  -- Sales Invoices
'		PI -- Purchase Invoices
'		D  -- Debit Notes
'		C  -- Credit Notes
'		P  -- Advance Payments
'		R  -- Advance Receipts
'*********************** Adjustment Type Details ******************************************
blnSalesInvoice = true

Dim iDocNo,iSno,iPayNo,iRecCount,sCrDrNoDesc
Dim sAmtReceived,sAmtReceiavble,sInvDate,sInvNo,sVouNo,dToAccount,dToAdjusted
Dim sCrDrNo,sCrDrType,iPlace

sQuery =" select a.CreatedReceivable,a.Narration,convert(Varchar(10),a.PartyInvoiceDate,103), "&_
		" a.AmountReceivable,a.AmountReceived,b.AmountReceived-a.AmountReceived,a.ReceivableNumber from Acc_T_Receivables a, Acc_T_CreatedReceivables b "&_
		" where a.OUDefinitionID='"&sOrgId&"' "&_
		" and b.ReceivableNumber=a.CreatedReceivable and a.AmountReceivable>a.AmountReceived   "&_
		" and a.PartyCode="&sParCode&" and a.VoucherDate <= convert(datetime,'"&sVouDate&"',103) "
		'" and A.ReceivableNumber NOT IN("&iITrNo&") "
'Response.Write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
iRecCount = objRs.RecordCount
Set objRs.ActiveConnection = nothing
iSno=1

Do While Not objRs.EOF
	blnSalesInvoice = false
	sInvNo = objRs(1)
	IF CStr(Trim(sInvNo)) = "&sNewNarr&" Then
		sQuery = "SELECT H.VoucherNumber, Convert(Varchar(10),H.VoucherDate,103) FROM Acc_T_Receivables R  "&_
				 "INNER JOIN Acc_T_VoucherHeader H ON R.TransactionNumber = H.TransactionNumber "&_
				 "WHERE R.CreatedReceivable = "&objRs(0)&" "

		'Response.Write sQuery
		objRs2.Open sQuery,Con
		IF Not objRs2.EOF Then
			sInvNo = "Sale Inv No:"&objRs2(0)& " Dt:"&objRs2(1)
		End IF
		objRs2.Close
	End IF

	sInvDate = objRs(2)
	iPayNo= objRs(6)
	iDocNo = objRs(0)
	sAmtReceiavble = Trim(objRs(3))
	sAmtReceived = Trim(objRs(4))
	dToAccount= Trim(objRs(5))
	'dToAdjusted = Trim(objRs(6))

	sAmtReceiavble = CDbl(sAmtReceiavble)
	sAmtReceived = CDbl(sAmtReceived)
	dToAccount= CDbl(dToAccount)
	dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
	
	sPrintAdjusted = "N"
	'Response.Write "<p><font color=red>dToAdjusted="&dToAdjusted
	IF CStr(iCrTransNo) <> 0 Then
		sDispClass = "ExcelDisplayCell"
		sExp = "//PayRec/Doc[@No="&iDocNo&"]"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			dToAdjusted = CDbl(dToAdjusted) + CDbl(TempNode.item(0).Attributes.getNamedItem("AmtToAdjust").Value)
			sAmtReceived = CDbl(sAmtReceived) - CDbl(TempNode.item(0).Attributes.getNamedItem("AmtToAdjust").Value)
			sDispClass = "ExcelDisplayCell2"
		End IF
	End IF
	
	IF CDbl(dToAdjusted) <> 0 Then
		sInvNo = Trim(Replace(sInvNo,"SALE INV NO:",""))
		iPlace = InStr(1,sInvNo,"Dt")
		sInvNo = Mid(sInvNo,1,iPlace-1)
		'Response.Write iPlace
		If InStr(1,sInvNo,"/") Then
			sInvNo = Replace(sInvNo,"/","")
		End IF

%>
			<tr>
				
				<% IF CStr(iCrTransNo) <> "0" Then %>
				<%		IF TempNode.length <> 0 Then 
							sPrintAdjusted = "Y"%>
							<td class="ExcelSerial" align="center"><%=iSno%><FONT color="red"><b>*</b></font></td>
				<%		Else%>
							<td class="ExcelSerial" align="center"><%=iSno%></td>
				<%		End IF
					Else
				%>
					<td class="ExcelSerial" align="center"><%=iSno%></td>
				<%End IF %>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=trim(sInvNo)%>Z<%=sPrintAdjusted%>" value="<%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?I">
				<%If sPrintAdjusted = "N" Then%>
					<td class="ExcelInputCell" align="right" width="10">
					<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=trim(sInvNo)%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
					<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
					<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right">
						<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=trim(sInvNo)%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
					</td>
				<%Else
					nAmtOutstanding = 0
					sAmtReceived = objRs(4)
					nAmtOutstanding = cdbl(sAmtReceiavble) - (cdbl(sAmtReceived)+ cdbl(dToAccount))
					
				%>
				<td class="ExcelInputCell" align="right" width="10">
					<%If sCheck = "A" Then
						If TempNode.item(0).Attributes.getNamedItem("Check").Value = "DEL" Then
							sAmtReceived = cdbl(sAmtReceived) - cdbl(TempNode.item(0).Attributes.getNamedItem("AmtToAdjust").Value)
							nAmtOutstanding = cdbl(sAmtReceiavble) - (cdbl(sAmtReceived)+ cdbl(dToAccount))%>
							<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=trim(sInvNo)%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
						<%Else%>
							<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=trim(sInvNo)%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
						<%End IF%>
					<%Else%>
						<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=trim(sInvNo)%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
					<%End IF%>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<Input type="Text" Name="TxtAdjust<%=sInvNo%>" Value="<%=FormatNumber(sAmtReceived,2,,,0)%>" style="text-align:right" size="10" class="FormElem" Readonly>
				</td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=trim(sInvNo)%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(nAmtOutstanding,2,,,0)%><%'=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
				<%End IF 'If sPrintAdjusted = "N" Then%>
				<td class="ExcelDisplayCell" align="right">
					<Input type="Text" Name="TxtToAdjust<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=trim(sInvNo)%>Z<%=sPrintAdjusted%>" Value="0.00" style="text-align:right" class="FormElem" Readonly onchange="DisplayAmt('<%=iDocNo&"Z"&iPayNo&"Z"&iSno&"Z"&trim(sInvNo)&"Z"& sPrintAdjusted%>')">
				</td>
            </tr>

		<%
			iSno=cint(iSno)+1
		End IF
		
		objRs.MoveNext
	Loop
	objRs.Close

	IF CStr(iAdvNo) <> "0" Then
		sQuery = "Select Distinct B.PayRecNo,B.BillAmt,B.BillAdjAmt,R.CreatedReceivable,R.Narration, "&_
				 "Convert(Varchar(10),R.PartyInvoiceDate,103) PartyInvoiceDate From "&_
				 "Acc_T_BillsClosingDetails B,Acc_T_Receivables R, " &_
				 "Acc_T_VoucherHeader H Where R.ReceivableNumber = B.PayRecNo And  "&_
				 "B.AdvCloseNo = "&iAdvNo&" And H.TransactionNumber = B.BillTrNo  "&_
				 "And H.BookCode = '05' "

		'Response.Write sQuery
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		iRecCount = objRs.RecordCount
		Set objRs.ActiveConnection = nothing
		Do While Not objRs.EOF
		blnSalesInvoice = false
			sInvNo = objRs("Narration")
			sInvDate = objRs("PartyInvoiceDate")
			iPayNo= objRs("PayRecNo")
			iDocNo = objRs("CreatedReceivable")
			sAmtReceiavble = Trim(objRs("BillAmt"))
			sAmtReceived = CDbl(objRs("BillAmt")) - CDbl(objRs("BillAdjAmt"))
			dToAccount= 0
			sAmtReceiavble = CDbl(sAmtReceiavble)
			sAmtReceived = CDbl(sAmtReceived)
			dToAccount= CDbl(dToAccount)
			dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)

			sInvNo = Trim(Replace(sInvNo,"SALE INV NO:",""))
			iPlace = InStr(1,sInvNo,"Dt")
			sInvNo = Mid(sInvNo,1,iPlace-1)
%>
			<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?I">
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)" checked></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
				<td class="ExcelDisplayCell" align="right"></td>
            </tr>

<%
			iSno=cint(iSno)+1
			objRs.MoveNext
		Loop
		objRs.Close
	End IF

	'Response.Write " blnSalesInvoice = "& blnSalesInvoice
	if blnSalesInvoice then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Sales Invoices Found !!</b></td>
		</tr>
	<%Else %>
	<tr>
		<td colspan="8" class="ExcelDisplayCell" align="Right"><b>Tatal Amount To Be Adjusted</b></td>
		<td class="ExcelDisplayCell" align="right">
			<Input type="Text" Name="TxtTotToAdjustedAmount" style="text-align:right" Value="0.00" Readonly Class="FormElemRead">
		</td>
	</tr>
	<%End If %>

<tr>
<td colspan="9" class="ExcelSerial"><b>
<%
	'IF CStr(sVouType) = "D" Then
		Response.Write "Debit Notes"
		sCrDrNoDesc = "DEBIT NOTE NO: "
	'Else
	'	Response.Write "Credit Notes"
	'	sCrDrNoDesc = "CREDIT NOTE NO: "
	'End IF
%>

</b></td>
</tr>
<%


Dim sCrDrVouDate
'===================================================================================
'Added On				:	18/11/2004
'Reason					:	To list out all related Credit/Debit notes related to party
'						:	I - Invoice
'						:	D - Debit Note
'						:	C - Credit Note
'===================================================================================
blnDebitNode = true
sQuery2 = "Select P.ReceivableNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyInvoiceDate,103),P.AmountReceivable, "&_
		  "P.AmountReceived,C.AmountReceived - P.AmountReceived,P.ReceivableNumber, Convert(Char,H.VoucherDate,103) "&_
		  "From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C Where "&_
		  "P.CreatedReceivable = 0 and   "&_
		  "P.AmountReceivable > P.AmountReceived and P.OUDefinitionID = '"&sOrgId&"'  "&_
		  "and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) "&_
		  "and P.TransactionNumber = H.TransactionNumber and C.ReceivableNumber = P.DRCreatedReceivable  "&_
		  "and C.AmountReceivable > C.AmountReceived and P.ReceivableNumber NOT IN("&iCNTrNo&") "

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery2
	.ActiveConnection = con
	.Open
end with

iRecCount = objRs.RecordCount + CDbl(iRecCount)


set objRs.ActiveConnection = nothing
'=====================================================================================


	Do While Not objRs.EOF
		blnDebitNode = false
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		iPayNo = objRs(0)
		sCrDrNo = objRs(6)
		sCrDrVouDate = objRs(7)

		sAmtReceiavble = Trim(objRs(3))
		sAmtReceived = Trim(objRs(4))
		dToAccount= Trim(objRs(5))

		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)

		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)


		IF CDbl(dToAdjusted) > 0 Then

%>
		<tr>
			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?D">
            <td class="ExcelSerial" align="center"><%=iSno%></td>
            <td class="ExcelInputCell" align="right" width="10">
            <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)"></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sCrDrVouDate%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right">
				<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
			</td>
			<td class="ExcelDisplayCell" align="right"></td>
       </tr>

<%
		iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
	Loop

objRs.Close

IF CStr(iAdvNo) <> "0" Then
	sQuery = "Select Distinct B.PayRecNo,B.BillAmt,B.BillAdjAmt,R.CreatedReceivable,R.Narration, "&_
			 "Convert(Varchar(10),R.PartyInvoiceDate,103) PartyInvoiceDate From "&_
			 "Acc_T_BillsClosingDetails B,Acc_T_Receivables R, " &_
			 "Acc_T_VoucherHeader H Where R.ReceivableNumber = B.PayRecNo And  "&_
			 "B.AdvCloseNo = "&iAdvNo&" And H.TransactionNumber = B.BillTrNo  "&_
			 "And H.BookCode = '06' "

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	iRecCount = objRs.RecordCount
	Set objRs.ActiveConnection = nothing
	Do While Not objRs.EOF
	blnDebitNode = false
		sInvNo = objRs("Narration")
		sInvDate = objRs("PartyInvoiceDate")
		iPayNo= objRs("PayRecNo")
		iDocNo = objRs("CreatedReceivable")
		sAmtReceiavble = Trim(objRs("BillAmt"))
		sAmtReceived = CDbl(objRs("BillAmt")) - CDbl(objRs("BillAdjAmt"))
		dToAccount= 0
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
%>
		<tr>
			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?D">
            <td class="ExcelSerial" align="center"><%=iSno%></td>
            <td class="ExcelInputCell" align="right" width="10">
            <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)" Checked></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sCrDrVouDate%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right">
				<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
			</td>
			<td class="ExcelDisplayCell" align="right"></td>
       </tr>

<%
		iSno=cint(iSno)+1
		objRs.MoveNext
	Loop
	objRs.Close
End IF
if blnDebitNode then
%>
	<tr>
		<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Debit Notes Found !!</b></td>
	</tr>
<%
end if

sCrDrNoDesc = "Purchase Biils "
%>

<tr><td colspan="9" class="ExcelSerial"><b>Purchase Bills </b></td></tr>
<%
'********************************** Sales Vouchers Over *****************************************
blnPurBills = true
sQuery3 = " Select a.CreatedPayablesNumber,a.Narration,convert(Varchar(10),a.PartyBillDate,103), "&_
		  " a.AmountPayable,a.AmountPaid,b.AmountPaid-a.AmountPaid,a.PayablesNumber,a.PartyBillNumber from Acc_T_Payables a, Acc_T_CreatedPayables b "&_
		  " where a.OUDefinitionID='"&sOrgId&"' and b.PayablesNumber=a.CreatedPayablesNumber  "&_
		  " and a.PartyCode="&sParCode&" and a.VoucherDate <= convert(datetime,'"&sVouDate&"',103) "
		  '" and a.PartyCode="&sParCode&" and a.AmountPayable>a.AmountPaid and a.VoucherDate <= convert(datetime,'"&sVouDate&"',103) "


'Response.write "<p><font color=red>"&sQuery3
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery3
	.ActiveConnection = con
	.Open
end with

iRecCount = objRs.RecordCount + CDbl(iRecCount)
Set objRs.ActiveConnection = nothing


'iSno=1
If not objRs.EOF then
	blnPurBills = false
	Do While Not objRs.EOF
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		iPayNo= objRs(6)
		sAmtReceiavble = Trim(objRs(3))
		sAmtReceived = Trim(objRs(4))
		dToAccount= Trim(objRs(5))
		'dToAdjusted = Trim(objRs(6))
		'nBillNo = Trim(objRs(7))

		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
		
		sPrintAdjusted = "N"
		
		IF CStr(iCrTransNo) <> 0 Then
			sDispClass = "ExcelDisplayCell"
			sExp = "//PayRec/Doc[@No="&iDocNo&"]"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.length <> 0 Then
				dToAdjusted = CDbl(dToAdjusted) + CDbl(TempNode.item(0).Attributes.getNamedItem("AmtToAdjust").Value)
				sAmtReceived = CDbl(sAmtReceived) - CDbl(TempNode.item(0).Attributes.getNamedItem("AmtToAdjust").Value)
				sDispClass = "ExcelDisplayCell2"
			End IF
		End IF

		IF CDbl(dToAdjusted) <> 0 Then
			sInvNo = Replace(sInvNo,"PUR INV No:","")
			iPlace = InStr(1,sInvNo,"Dt")
			sInvNo = Trim(Mid(sInvNo,1,iPlace-1))
			'Response.Write "<p><font color=red>sInvNo="&sInvNo
			
		%>

		<tr>

			
		<% IF CStr(iCrTransNo) <> "0" Then %>
		<%		IF TempNode.length <> 0 Then 
					sPrintAdjusted = "Y"%>
					<td class="ExcelSerial" align="center"><%=iSno%><FONT color="red"><b>*</b></font></td>
		<%		Else%>
					<td class="ExcelSerial" align="center"><%=iSno%></td>
		<%		End IF
			Else
		%>
			<td class="ExcelSerial" align="center"><%=iSno%></td>
		<%End IF %>
		<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?PI">
		<%If sPrintAdjusted = "N" Then%>
            <td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" ></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>

            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td-->
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right">
				<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
			</td>
		<%Else
			nAmtOutstanding = 0
			'IF TempNode.length <> 0 Then 
			'	sAmtReceiavble = TempNode.item(0).Attributes.getNamedItem("TransAmount").value
			'	sAmtReceived = TempNode.item(0).Attributes.getNamedItem("AmtToAdjust").value
			'	dToAccount = TempNode.item(0).Attributes.getNamedItem("AmtToAccount").value
			'	nAmtOutstanding = cdbl(sAmtReceiavble) - (cdbl(sAmtReceived)+ cdbl(dToAccount))
			'End IF 
			sAmtReceived = objRs(4)
			nAmtOutstanding = cdbl(sAmtReceiavble) - (cdbl(sAmtReceived)+ cdbl(dToAccount))
			%>
			<td class="ExcelInputCell" align="right" width="10">
				<%If sCheck = "A" Then
					If TempNode.item(0).Attributes.getNamedItem("Check").Value = "DEL" Then
						sAmtReceived = cdbl(sAmtReceived) - cdbl(TempNode.item(0).Attributes.getNamedItem("AmtToAdjust").Value)
						nAmtOutstanding = cdbl(sAmtReceiavble) - (cdbl(sAmtReceived)+ cdbl(dToAccount))%>
						<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
					<%Else%>
						<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
					<%End IF%>
				<%Else%>
					<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)" Disabled></td>
				<%End IF%>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>

            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right">
				<Input type="Text" Name="TxtAdjust<%=sInvNo%>" Value="<%=FormatNumber(sAmtReceived,2,,,0)%>" style="text-align:right" size="10" class="FormElem" Readonly>
			</td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right">
				<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(nAmtOutstanding,2,,,0)%><%'=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
			</td>
		<%End IF 'If sPrintAdjusted = "N" Then%>
			<td class="ExcelDisplayCell" align="right">
				<Input type="Text" Name="TxtToAdjust<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" Value="0.00" style="text-align:right" class="FormElem" Readonly onchange="DisplayAmt('<%=iDocNo&"Z"&iPayNo&"Z"&iSno&"Z"&sInvNo&"Z"& sPrintAdjusted%>')">
			</td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td-->
       </tr>

<%
			iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
	Loop
End IF
objRs.Close
%>

<%
IF CStr(iAdvNo) <> "0" Then
	sQuery = "Select Distinct B.PayRecNo,B.BillAmt,B.BillAdjAmt,R.CreatedPayablesNumber,R.Narration, "&_
			 "Convert(Varchar(10),R.PartyBillDate,103) PartyBillDate From "&_
			 "Acc_T_BillsClosingDetails B,Acc_T_Payables R, " &_
			 "Acc_T_VoucherHeader H Where R.PayablesNumber = B.PayRecNo And  "&_
			 "B.AdvCloseNo = "&iAdvNo&" And H.TransactionNumber = B.BillTrNo  "&_
			 "And H.BookCode = '04' "

	'Response.Write sQuery

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	iRecCount = objRs.RecordCount
	Set objRs.ActiveConnection = nothing
	Do While Not objRs.EOF
		blnPurBills = false
		sInvNo = objRs("Narration")
		sInvDate = objRs("PartyBillDate")
		iPayNo= objRs("PayRecNo")
		iDocNo = objRs("CreatedPayablesNumber")
		sAmtReceiavble = Trim(objRs("BillAmt"))
		sAmtReceived = CDbl(objRs("BillAmt")) - CDbl(objRs("BillAdjAmt"))
		dToAccount= 0
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)

		sInvNo = Replace(sInvNo,"PUR INV No:","")
		iPlace = InStr(1,sInvNo,"Dt")
		sInvNo = Trim(Mid(sInvNo,1,iPlace-1))
%>

		<tr>
			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?PI">
			<td class="ExcelSerial" align="center"><%=iSno%></td>
		    <td class="ExcelInputCell" align="right" width="10">
            <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)" Checked></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td-->
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right">
				<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
			</td>
			<td class="ExcelDisplayCell" align="right"></td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td-->
       </tr>

<%
		iSno=cint(iSno)+1
		objRs.MoveNext
	Loop
objRs.Close
End IF
if blnPurBills then
%>
	<tr>
		<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Purchse Bills Found !!</b></td>
	</tr>
<%
Else
%>
<tr>
	<td colspan="8" class="ExcelDisplayCell" align="Right"><b>Tatal Amount To Be Adjusted</b></td>
	<td class="ExcelDisplayCell" align="right">
		<Input type="Text" Name="TxtTotToAdjustedAmount" style="text-align:right" Value="0.00" Readonly Class="FormElemRead">
	</td>
</tr>
<%
end if

sCrDrNoDesc = "Credit Notes "
%>
<tr><td colspan="9" class="ExcelSerial"><b>Credit Notes </b></td></tr>
<%
blnCreditNode = true
'**************************** Purchase Invoice Over **************************************
sQuery4 = "Select P.PayablesNumber,H.VoucherNumber,Convert(Varchar(10),P.PartyBillDate,103), "&_
		  "P.AmountPayable,P.AmountPaid,C.AmountPaid - P.AmountPaid ,P.PayablesNumber, Convert(Char,H.VoucherDate,103) "&_
		  "From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "&_
		  "P.CreatedPayablesNumber = 0 and P.Narration is Null and P.AmountPayable > P.AmountPaid "&_
		  "and P.OUDefinitionID = '"&sOrgId&"' and P.PartyCode = "&sParCode&" and "&_
		  "P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) and P.TransactionNumber = "&_
		  "H.TransactionNumber and C.PayablesNumber = P.CrCreatedPayable "&_
		  "And C.AmountPayable > C.AmountPaid and P.PayablesNumber NOT IN("&iCNTrNo&") "


'Response.Write sQuery4

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery4
	.ActiveConnection = con
	.Open
end with

iRecCount = objRs.RecordCount + CDbl(iRecCount)


set objRs.ActiveConnection = nothing
'=====================================================================================
Do While Not objRs.EOF
	blnCreditNode = false
	iDocNo = objRs(0)
	sInvNo = objRs(1)
	sInvDate = objRs(2)
	iPayNo = objRs(0)
	sCrDrNo = objRs(6)
	sCrDrVouDate = objRs(7)

	sAmtReceiavble = Trim(objRs(3))
	sAmtReceived = Trim(objRs(4))
	dToAccount= Trim(objRs(5))

	sAmtReceiavble = CDbl(sAmtReceiavble)
	sAmtReceived = CDbl(sAmtReceived)
	dToAccount= CDbl(dToAccount)
	dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
	IF CDbl(dToAdjusted) > 0 Then

%>
		 <tr>
			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?C">
		     <td class="ExcelSerial" align="center"><%=iSno%></td>
		     <td class="ExcelInputCell" align="right" width="10">
		     <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)"></td>
		     <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
		     <td class="ExcelDisplayCell"><p align="center"><%=sCrDrVouDate%></td>
		     <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
		     <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
		     <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
		     <td class="ExcelDisplayCell" align="right">
		 		<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
		 	</td>
		 	<td class="ExcelDisplayCell" align="right"></td>
		</tr>

<%
			iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
	Loop

objRs.Close

IF CStr(iAdvNo) <> "0" Then
	sQuery = "Select Distinct B.PayRecNo,B.BillAmt,B.BillAdjAmt,R.CreatedPayablesNumber,H.VoucherNumber, "&_
			 "Convert(Varchar(10),H.VoucherDate,103) VoucherDate From "&_
			 "Acc_T_BillsClosingDetails B,Acc_T_Payables R, " &_
			 "Acc_T_VoucherHeader H Where R.PayablesNumber = B.PayRecNo And  "&_
			 "B.AdvCloseNo = "&iAdvNo&" And H.TransactionNumber = B.BillTrNo  "&_
			 "And H.BookCode = '07' "

	'Response.Write sQuery

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	iRecCount = objRs.RecordCount
	Set objRs.ActiveConnection = nothing
	Do While Not objRs.EOF
		blnCreditNode = false
		sInvNo = objRs("VoucherNumber")
		sInvDate = objRs("VoucherDate")
		iPayNo= objRs("PayRecNo")
		iDocNo = objRs("CreatedPayablesNumber")
		sAmtReceiavble = Trim(objRs("BillAmt"))
		sAmtReceived = CDbl(objRs("BillAmt")) - CDbl(objRs("BillAdjAmt"))
		dToAccount= 0
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)


%>

		<tr>
			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?PI">
			<td class="ExcelSerial" align="center"><%=iSno%></td>
		    <td class="ExcelInputCell" align="right" width="10">
            <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)" Checked></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td-->
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right">
				<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
			</td>
			<td class="ExcelDisplayCell" align="right"></td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td-->
       </tr>

<%
		iSno=cint(iSno)+1
		objRs.MoveNext
	Loop
objRs.Close
End IF

if blnCreditNode then
%>
	<tr>
		<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Credit Notes Found !!</b></td>
	</tr>
<%
end if


'************************ Credit Note Over ************************************************

'Response.Write iCrTransNo
IF CStr(iCrTransNo) <> "0" Then
'************************ Advance Payments/Receipts ****************************************
	IF CStr(sParType) = "CR" Then
	sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),isnull(a.AdvanceAdjusted,0), "&_
			 "isnull(a.AdvancePaid,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
			 "b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
			 "A.AdvanceNumber,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
			 "from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
			 "a.OUDefinitionID='"&sOrgId&"' and a.PartyCode="&sParCode&" and "&_
			 "isnull(a.AdvancePaid,0)>isnull(a.AdvanceAdjusted,0) and "&_
			 "b.TransactionNumber = a.TransactionNumber "&_
			 "and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
			 "A.CreatedAdvanceNo = C.CreatedAdvanceNo  and A.AdvanceNumber NOT IN("&iAdvPayTrNo&") "
	Else
	sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),isnull(a.AdvanceAdjusted,0), "&_
			 "isnull(a.AdvanceReceived,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
			 "b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
			 "A.AdvanceNumber,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
			 "from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
			 "a.OUDefinitionID='"&sOrgId&"' and a.PartyCode="&sParCode&" and "&_
			 "isnull(a.AdvanceReceived,0)>isnull(a.AdvanceAdjusted,0) and "&_
			 "b.TransactionNumber = a.TransactionNumber "&_
			 "and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
			 "A.CreatedAdvanceNo = C.CreatedAdvanceNo and A.AdvanceNumber NOT IN("&iAdvRecTrNo&") "
	End IF

	'Response.Write sQuery &"<BR>"

	IF CStr(sParType) = "CR" Then
		blnAdvancePay = true
		sCrDrType = "P"
		sCrDrNoDesc = "Advance Payments  "
%>
		<!--<tr>
			<td class="ExcelHeaderCell" align="left" colspan="9"><b>Advance Payments</b></td>
		</tr>-->
<%	Else
	blnAdvanceRec = true
		sCrDrType = "R"
		sCrDrNoDesc = "Advance Receipts  "
%>
		<tr>
			<td class="ExcelHeaderCell" align="left" colspan="9"><b>Advance Receipts</b></td>
		</tr>
<%
	End IF

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	'iSno=1
	Do While Not objRs.EOF
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		sAmtReceived = objRs(3)
		sAmtReceiavble = objRs(4)
		'iCrDocNo = objRs(5)
		iPayNo = objRs(12)
		dToAccount = Objrs(13)
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
		IF CDbl(dToAdjusted) <> 0 Then
			blnAdvancePay = false
			blnAdvanceRec = false
			sPrintAdjusted = ""
%>
			<!--<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>">
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" class="FormElem" onClick="UpdateAmt(this)"></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
				<td class="ExcelDisplayCell" align="right">
					<Input type="Text" Name="TxtToAdjust<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>Z<%=sInvNo%>Z<%=sPrintAdjusted%>" Value="0.00" style="text-align:right" class="FormElem" Readonly >
				</td>
			</tr>-->

<%			iSno=cint(iSno)+1
		End IF
	objRs.MoveNext
	Loop
	objRs.Close

'	Response.Write "sCrDrType = "&sCrDrType
'	Response.Write  blnAdvancePay
'	Response.Write blnAdvanceRec
	if trim(sCrDrType)="P" and blnAdvancePay then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Advance Payments Found !!</b></td>
		</tr>
	<%
	elseif trim(sCrDrType)="R" and blnAdvanceRec then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Advance Receipts Found !!</b></td>
		</tr>
	<%
	end if

	IF CStr(iCrTransNo) <> "0" Then
%>
		<tr>
			<td class="ExcelDisplayCell" colspan="9"><b><FONT color="red">*</Font> Denotes Bill(s) Adjusted to this Voucher </b></td>
		</tr>

<%
	End IF
'End IF
Else
	'Response.Write "==========="
	sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),isnull(a.AdvanceAdjusted,0), "&_
			 "isnull(a.AdvancePaid,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
			 "b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
			 "A.AdvanceNumber,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
			 "from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
			 "a.OUDefinitionID='"&sOrgId&"' and a.PartyCode="&sParCode&" and "&_
			 "isnull(a.AdvancePaid,0)>isnull(a.AdvanceAdjusted,0) and "&_
			 "b.TransactionNumber = a.TransactionNumber "&_
			 "and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
			 "A.CreatedAdvanceNo = C.CreatedAdvanceNo  and A.AdvanceNumber NOT IN("&iAdvPayTrNo&")  Order By b.VoucherDate"

	sQuery2 = "select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103),isnull(a.AdvanceAdjusted,0), "&_
			 "isnull(a.AdvanceReceived,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
			 "b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
			 "A.AdvanceNumber,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
			 "from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
			 "a.OUDefinitionID='"&sOrgId&"' and a.PartyCode="&sParCode&" and "&_
			 "isnull(a.AdvanceReceived,0)>isnull(a.AdvanceAdjusted,0) and "&_
			 "b.TransactionNumber = a.TransactionNumber "&_
			 "and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
			 "A.CreatedAdvanceNo = C.CreatedAdvanceNo and A.AdvanceNumber NOT IN("&iAdvRecTrNo&")  Order By b.VoucherDate "
%>
	<tr>
		<td class="ExcelHeaderCell" align="left" colspan="9"><b>Advance Payments</b></td>
	</tr>
<%
	sCrDrType = "P"
	sCrDrNoDesc = "Advance Payments  "
	blnAdvancePay = true
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	'iSno=1
	Do While Not objRs.EOF
		blnAdvancePay = false
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		sAmtReceived = objRs(3)
		sAmtReceiavble = objRs(4)
		'iCrDocNo = objRs(5)
		iPayNo = objRs(12)
		dToAccount = Objrs(13)
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
		IF CDbl(dToAdjusted) <> 0 Then

%>
			<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>">
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)"></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
				<td class="ExcelDisplayCell" align="right"></td>
			</tr>

<%			iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
		Loop
		objRs.Close

		IF CStr(iAdvNo) <> "0" Then
			sQuery = "Select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103) "&_
					 "VoucherDate,D.BillAdjAmt, D.BillAmt,B.TransactionType,A.CreatedTransNo, "&_
					 "A.AdvanceNumber From Acc_T_AdvancePayments A, Acc_T_VoucherHeader b, "&_
					 "Acc_T_CreatedAdvances C,Acc_T_BillsClosingDetails D where a.OUDefinitionID='"&sOrgId&"' and "&_
					 "A.PartyCode="&sParCode&" and B.TransactionNumber = A.TransactionNumber and  "&_
					 "A.CreatedAdvanceNo = C.CreatedAdvanceNo And D.BillTrNo = C.CreatedAdvanceNo And "&_
					 "D.AdvCloseNo = "&iAdvNo&" And D.BillType = 'P' Order By b.VoucherDate "

			'Response.Write sQuery
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs.ActiveConnection = nothing
			'iSno=1
			Do While Not objRs.EOF
				blnAdvance = false
			iDocNo = objRs("TransactionNumber")
			sInvNo = objRs("VoucherNumber")
			sInvDate = objRs("VoucherDate")
			sAmtReceived = objRs("BillAdjAmt")
			sAmtReceiavble = objRs("BillAmt")
			'iCrDocNo = objRs(5)
			iPayNo = objRs("AdvanceNumber")
			dToAccount = 0
			sAmtReceiavble = CDbl(sAmtReceiavble)
			sAmtReceived = CDbl(sAmtReceived)
			dToAccount= CDbl(dToAccount)
			'dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
			dToAdjusted = sAmtReceived
			IF CDbl(dToAdjusted) <> 0 Then

%>
				<tr>
					<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>">
					<td class="ExcelSerial" align="center"><%=iSno%></td>
					<td class="ExcelInputCell" align="right" width="10">
					<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)" Checked></td>
					<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
					<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
					<td class="ExcelDisplayCell" align="right">
						<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
					</td>
					<td class="ExcelDisplayCell" align="right"></td>
				</tr>

<%				iSno=cint(iSno)+1
			End IF
			objRs.MoveNext
		Loop
		objRs.Close

	End IF

	if blnAdvancePay then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Advance Payments Found !!</b></td>
		</tr>
	<%
	end if

%>
	<tr>
		<td class="ExcelHeaderCell" align="left" colspan="9"><b>Advance Receipts</b></td>
	</tr>
<%
	sCrDrType = "R"
	sCrDrNoDesc = "Advance Receipts  "
	blnAdvanceRec = true
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery2
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	'iSno=1
	Do While Not objRs.EOF
		blnAdvanceRec =false
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		sAmtReceived = objRs(3)
		sAmtReceiavble = objRs(4)
		'iCrDocNo = objRs(5)
		iPayNo = objRs(12)
		dToAccount = Objrs(13)
		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)
		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
		IF CDbl(dToAdjusted) <> 0 Then

%>
			<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>">
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)" ></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right">
					<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
				</td>
				<td class="ExcelDisplayCell" align="right"></td>
			</tr>

<%			iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
		Loop
		objRs.Close

		IF CStr(iAdvNo) <> "0" Then
			sQuery = "Select a.TransactionNumber,b.VoucherNumber,convert(Varchar(10),b.VoucherDate,103) "&_
					 "VoucherDate,D.BillAdjAmt, D.BillAmt,B.TransactionType,A.CreatedTransNo, "&_
					 "A.AdvanceNumber From Acc_T_AdvancePayments A, Acc_T_VoucherHeader b, "&_
					 "Acc_T_CreatedAdvances C,Acc_T_BillsClosingDetails D where a.OUDefinitionID='"&sOrgId&"' and "&_
					 "A.PartyCode="&sParCode&" and B.TransactionNumber = A.TransactionNumber and  "&_
					 "A.CreatedAdvanceNo = C.CreatedAdvanceNo And D.BillTrNo = C.CreatedAdvanceNo And "&_
					 "D.AdvCloseNo = "&iAdvNo&" And D.BillType = 'R' Order By b.VoucherDate "


			'Response.Write sQuery
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			Set objRs.ActiveConnection = nothing
			Do While Not objRs.EOF
				blnAdvanceRec = false
				iDocNo = objRs("TransactionNumber")
				sInvNo = objRs("VoucherNumber")
				sInvDate = objRs("VoucherDate")
				sAmtReceived = objRs("BillAdjAmt")
				sAmtReceiavble = objRs("BillAmt")
				'iCrDocNo = objRs(5)
				iPayNo = objRs("AdvanceNumber")
				dToAccount = 0
				sAmtReceiavble = CDbl(sAmtReceiavble)
				sAmtReceived = CDbl(sAmtReceived)
				dToAccount= CDbl(dToAccount)
				'dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
				dToAdjusted = sAmtReceived
				IF CDbl(dToAdjusted) <> 0 Then

%>
					<tr>
						<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>">
						<td class="ExcelSerial" align="center"><%=iSno%></td>
						<td class="ExcelInputCell" align="right" width="10">
						<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem" onClick="UpdateAmt(this)" Checked></td>
						<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
						<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
						<td class="ExcelDisplayCell" align="right">
							<input type="text" name="txtToAdk<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" size="13" class="FormElemRead" readonly Value="<%=FormatNumber(dToAdjusted,2,,,0)%>" style="text-align:right">
						</td>
						<td class="ExcelDisplayCell" align="right"></td>
					</tr>

<%				iSno=cint(iSno)+1
			End IF
			objRs.MoveNext
		Loop
		objRs.Close
		End IF

	if blnAdvancePay then
	%>
		<tr>
			<td colspan="9" class="ExcelDisplayCell" align=Center ><b>No Advance Receipts Found !!</b></td>
		</tr>
	<%
	end if

	End IF
End IF



 %>
  </table>                                              <Input type="hidden" name="hRecCount" value="<%=iRecCount%>">
												</div>
								</td>
								<td align="center">
								</td>
							</tr>

							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>

							<tr>
								<td align="center">
			<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
                                                            <%'IF CStr(iCtr) <> "0" Then %>
																	<input type="button" value="Done" name="B8" class="ActionButton" onClick="finaldone()" >
																	<!--input type="button" value="Cancel" name="B8" class="ActionButton" onClick="finalcancel()" -->
																	<!--input type="reset" value="Reset" name="B9" class="ActionButton"-->
															<%'Else%>
																<input type="button" value="Cancel" name="B8" class="ActionButton" onClick="finalcancel()">
															<%'End IF%>
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
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

<%'Else%>
	<!--
	<tr>
	<td class="ExcelDisplayCell"><b>Party is Not Selected in this Voucher Adjustment is Not Possible!! </b></td>
	<tr>
		<td valign="top">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td valign="middle" class="ActionCell">
		            <p align="center">
		            -->
		                <!--input type="button" value="Done" name="B8" class="ActionButton" onClick="finaldone()" -->
					<!--<input type="button" value="Cancel" name="B8" class="ActionButton" onClick="finalcancel()">-->
						<!--input type="reset" value="Reset" name="B9" class="ActionButton"-->
					<!--
				</td>
			</tr>
		</table>
		</td>
	</tr>
	-->
<%'End IF%>

</table>
</form>
</BODY>
</HTML>
