<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MsiVouEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Aug 20, 2004
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/Salpopulate.asp"-->

<%
dim sOrgId,sBookCode,sVouType,sOrgName,sBookName,oDOM
dim sQuery,objRs,bOtherUnits,iBookAccHead,dTransLimit
dim dOpeningBal,iTransNo,sPurVouNarr
Dim sPayTo,sVouNo,sVouDate,sVouAmount,sTransTy

dim sCreatedMiscPymtNo,sPaymentAgainst,sPayRefNo,iInvNo,sInvCode
dim iRcptCode,sGRNAgainstStr,sReceiptRouteStr,sReceiptCode,sItemType
Dim sReceiptRouting,iGRNNo,iRcptNo,iInspNo,sGrnCode,sGrnAgainst
Dim sRcptStr,sInspCode,saTemp,sUserID,sCrDrIndi,sAccUnit,sAccUnitName,sDispVal
Dim sFinPeriod,sFinFrom,sFinTo,sArrFin
Dim sPartyType,sParSubType,sPartyCode
set objRs = Server.CreateObject("ADODB.Recordset")

sOrgId=Request("OrgID")
sOrgName=Request("OrgName")
iTransNo = Request("TransNo")
sVouType="C"
sUserID = getUserID()

sFinPeriod = Session("FinPeriod")
sArrFin = Split(sFinPeriod,":")
sFinFrom = "01/04/"&sArrFin(0)
sFinTo = "31/03/"&sArrFin(1)

sQuery = "Select PayToRecdFrom,CreatedMiscPymtNo,Convert(Char,VoucherDate,103), "&_
		 "VoucherAmount,isNull(PaymentAgainst,''),isNull(ReferenceNo,0),CrDrIndication,TransactionType,PartyType,PartySubType,PartyCode From Acc_T_MiscPymtRequestHeader  "&_
		 "Where MiscTransNo = "&iTransNo&" "
		 


With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With

Set objRs.ActiveConnection = Nothing

sCreatedMiscPymtNo = ""

IF Not objRs.EOF Then
	sPayTo = objRs(0)
	sVouNo = objRs(1)
	sVouDate = objRs(2)
	sVouAmount = objRs(3)
	
	sCreatedMiscPymtNo = objRs(1)
	sPaymentAgainst = UCase(objRs(4))
	sPayRefNo = objRs(5)
	sCrDrIndi = objRs(6)
	sTransTy = objRs(7)
	sPartyType = objRs(8)
	sParSubType = objRs(9)
	sPartyCode = objRs(10)
End IF
objRs.Close

'Response.Write sTransTy

IF CStr(sTransTy) = "CAP" Then
	sVouType = "C"
Else 'Cash Receipt Only in Sales Invoices.
	sVouType = "D"
End IF


IF CStr(sCrDrIndi) = "D" Then
	'added by kalai selvi on 17/09/2004
	'****
	if trim(sPaymentAgainst) = "I" then
		sQuery = "Select distinct isNull(InvoiceNumber,0),(isnull(SuppInvoiceNo,InvoiceCode) + '--' + convert(varchar,InvoiceDate,103)) from RCV_T_InvoiceHeader where isNull(InvoiceNumber,0) = " & sPayRefNo & " "
		with objRs
			.ActiveConnection = con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open 
		end with
														
		set objRs.ActiveConnection = nothing

		iInvNo = 0
		sInvCode = ""
		if not objRs.EOF then
			iInvNo		= objRs(0)
			sInvCode	= objRs(1)
		end if 	'if not objRs.EOF then
		objRs.Close 
	end if 'if trim(sPaymentAgainst) = "I" then
	
	if trim(sPaymentAgainst) = "R" then
		sQuery = " Select Distinct GRNNumber,isnull(ReceiptNumber,0) RcptNum ,isnull(InvoiceNumber,0),isnull(InspectionNumber,0)," &_
				" GRNCode = (Select (GRNCode + '--' + convert(varchar,GRNDate,103)) AS GRN from Rcv_T_GateReceiptHeader where GRNNumber = R.GRNNumber)," &_
				" GRNAgainst = (Select ReceiptAgainst from Rcv_T_GateReceiptHeader where GRNNumber = R.GRNNumber), " &_
				" RCPTCode = isnull((Select (ReceiptCode + '--' + convert(varchar,ReceiptDate,103) + '|' + isnull(ReceiptRouting,0)) AS Rcpt from RCV_T_ActualReceiptHeader Where ReceiptNumber = R.ReceiptNumber),0)," &_
				" InvCode = isnull((Select (isnull(SuppInvoiceNo,InvoiceCode) + '--' + convert(varchar,SuppInvoiceDate,103)) AS INV from RCV_T_InvoiceHeader where InvoiceNumber = R.InvoiceNumber),0)," &_
				" InspCode = isnull((Select Distinct InspectionNumber from RCV_T_PurchInspectionHeader where InspectionNumber = R.InspectionNumber),0), " &_
				" ItemType = isNull((Select Distinct IT.ItemtypeID from RCV_T_GRNItemDetails IT where  IT.GrnNumber = R.GRNNumber),'')" & _
				" from PUR_T_RefferenceNumberDet R where isnull(ReceiptNumber,0) = " & sPayRefNo & ""

		with objRs
			.ActiveConnection = con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open 
		end with
														
		set objRs.ActiveConnection = nothing


		if not objRs.EOF then
			iGRNNo = objRs(0)
			iRcptNo = objRs(1)
			iInspNo = objRs(3)
			sGrnCode = objRs(4)
			sGrnAgainst = objRs(5)
			sRcptStr = objRs(6)
													
			sInspCode = objRs(8)
																	
			sItemType = objRs(9)
			
			sGRNAgainstStr = getReceiptType(sGrnAgainst)
			sReceiptCode = ""
			sReceiptRouting = ""
			
			if sRcptStr <> "" and sRcptStr <> "0" then
				saTemp = Split(sRcptStr,"|")
				sReceiptCode  = saTemp(0)
				sReceiptRouting = saTemp(1)
			End if
			
			If trim(sReceiptRouting) = "" or trim(sReceiptRouting) = "0" Then
				sReceiptRouteStr = "--"
			Else
				sReceiptRouteStr = getReceiptRoute(sReceiptRouting)
			End if 
		end if 	'if not objRs.EOF then
		objRs.Close 
	end if 'if trim(sPaymentAgainst) = "I" then
End IF
	
	
																
'****

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

sQuery = "Select D.VoucherNarration,H.PayToRecdFrom From Acc_T_MiscPaymentReqDetails D,  "&_
		 "Acc_T_MiscPymtRequestHeader H Where D.MiscTransNo = "&iTransNo&" and H.MiscTransNo = "&iTransNo&" "
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With

Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	sPurVouNarr = objRs(0)
	'sPayTo = objRs(1)
End IF
objRs.Close
sPurVouNarr = Trim(sPurVouNarr)

'************* LMC Only Getting Accounting Unit From an XML File For YPD Accounting ********
'sAccUnit = GetAccountUnit(sOrgId)
'IF CStr(sOrgId) <> CStr(sAccUnit) Then
'	sQuery = "Select OrgUnitDescription From DCS_OrganizationUnitDefinitions Where OUDefinitionID = '"&sAccUnit&"' "
'	objRs.Open sQuery,Con
'	IF Not objRs.EOF Then
'		sAccUnitName = objRs(0)
'	End IF
'	objRs.Close
'Else
'	sAccUnit = sOrgId
'	sAccUnitName = sOrgName
'End IF
'*********** Added by Manohar Prabhu.R on 14/06/2005 ***************************************

	sAccUnit = sOrgId
	sAccUnitName = sOrgName


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script src="../../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script src="../../scripts/ExcelFunctions.js"></script>
<!--XML ISLAND FOR VOUCHER DATA -->
<script type="application/xml" data-itms-xml-island="1" id="VoucherData"><voucher UnitNo="<%=sAccUnit%>" UnitName="<%=sAccUnitName%>" BookNo="" BookName="" CRDR="<%=sVouType%>" VouDate="" BookAcchead="" Approver=""/></script>
<!--XML ISLAND FOR ENTRY DATA -->
<script type="application/xml" data-itms-xml-island="1" id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="" TDSElgi="0" TdsPercentage="0" /></script>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="TDSData"  ><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>
<script type="application/xml" data-itms-xml-island="1" ID="UnitBookData">
<Book/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempXMLData"><Root></Root></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script src="../../scripts/MiscVoucherEntryCompat.js"></script>
<script>
window.ITMSMiscVoucherEntry.install({
	bookCode: "01",
	moduleCode: "CA",
	bank: false,
	enableTds: true,
	transLimit: "<%=dTransLimit%>"
});
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init();setParty();">
<form method="POST" name="formname" action="VouMsiGenerate.asp">
<input type="hidden" name="hVouCode" value="01">
<input type="hidden" name="hVouCRDR" value="<%=sVouType%>">
<input type="hidden" name="hVouName" value="CA">
<input type="hidden" name="hOrgId" value="<%=sAccUnit%>">
<input type="hidden" name="hOrgName" value="<%=sAccUnitName%>">
<input type="hidden" name="hBookcode" value="<%=sBookCode%>">
<input type="hidden" name="hOtherUnitFlag" value="1">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hTdsElgi" value="0">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hUpdate" value="N">
<input type="hidden" name="hTdsNew" value="N">
<input type="hidden" name="hTdsAmt" value="">
<input type="hidden" name="hAmendTy" value="N">
<input type="hidden" name="hFinFrom" value="<%=sFinFrom%>" />
<input type="hidden" name="hFinTo" value="<%=sFinTo%>" />
<input type="hidden" name="hParType" value="<%=sPartyType%>" />
<input type="hidden" name="hParSubType" value="<%=sParSubType%>" />
<input type="hidden" name="hParCode" value="<%=sPartyCode%>" />

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		 Miscellaneous Voucher 
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Voucher Details
											</td>
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
                            
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
                                                            <table border="0" cellspacing="0" cellpadding="0">

                                                        <tr>
                                                    <td class="FieldCell" colspan="2">
                                                      <table border="0" width="100%" cellspacing="1" class="TableOutlineOnly">
                                                        <tr>
                                                          <td class="MiddlePack" colspan="6"></td>
                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="90">Voucher
                                                            Date</td>
                                                          <td class="FieldCellSub" width="125">
                                                          <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>

                                                          </td>
                                                          <td class="FieldCellSub" width="80">Entry
                                                      Number</td>
                                                          <td class="FieldCellSub"><span class="DataOnly" id="spEntryNo">1&nbsp;</span></td>
                                                          <td class="FieldCellSub" width="100">
                                                    </td>
                                                          <td class="FieldCellSub">
   <span class="DataOnly">
                                                            <%
                                                             'dOpeningBal =GetDayOpening(sOrgId,iBookAccHead,FormatDate(date+1))
                                                             'dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
                                                             'if dOpeningBal<0 then
															'	Response.Write dOpeningBal*-1 &"&nbsp;Cr"
															 'else
															'	Response.Write dOpeningBal &"&nbsp;Dr"
															 'end if
                                                            %></span>
   &nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="90">Entry
                                                            Type</td>
                                                          <td class="FieldCellSub" colspan="3">
											<%
											'Response.Write sVouType &"========"
											if sVouType="C" then%>
                                                            <input type=radio name="selCRDR" value="C" disabled>Receipts
                                                            <input type=radio name="selCRDR" value="D" checked>Payments
                                                            <%else%>
                                                            <input type=radio name="selCRDR" value="C" checked>Receipts
                                                            <input type=radio name="selCRDR" value="D" disabled >Payments&nbsp;
											<%end if%>
                                                            </td>
                                                          <td class="FieldCellSub" width="100">
                                                      &nbsp;</td>
                                                          <td class="FieldCellSub"><span class="DataOnly">
                                                            <%
                                                            ' dOpeningBal =GetDayOpeningCreated(sOrgId,iBookAccHead,FormatDate(date+1))
                                                            ' dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
                                                            ' if dOpeningBal<0 then
															'	Response.Write dOpeningBal*-1 &"&nbsp;Cr"
															' else
															'	Response.Write dOpeningBal &"&nbsp;Dr"
															' end if
                                                            %></span>
                                                            &nbsp;</td>
                                                        </tr>
                                                        <tr>
														  <td class="FieldCell" width="90">Accounting Unit </td>
														  <td class="FieldCellSub">
															<%Response.Write(sAccUnitName)%>
														  </td>	
														
                                                          <td class="FieldCellSub" width="90">Select 
                                                            Book</td>
                                                          <td class="FieldCellSub" colspan="3">
                                                          <select size="1" name="selBook" class="FormElem">
																<option value="S">Select Book</option>
															<%
																sQuery = "select BookNumber,BookName,isnull(BookAccountHead,0),OtherUnitTransaction from "&_
																		 "vwOrgBookNames where OUDefinitionID = '" & sAccUnit & "' and BookCode='01' "&_
																		 "and BookAccountHead is not null "
		
																with objRs
																	.CursorLocation = 3
																	.CursorType = 3
																	.Source = sQuery
																	.ActiveConnection = con
																	.Open
																end with
																set objRs.ActiveConnection = nothing
																while not objRs.EOF
															%>
																	<Option value="<%=objRs(0)%>?<%=objRs(2)%>"><%=objRs(1)%></Option>
															<%
																objRs.MoveNext
																Wend
																objRs.Close
															%>
                                                          </td>
                                                        </tr>
                                                        
                                                        <tr>
                                                          <td class="MiddlePack" colspan="6"></td>
                                                        </tr>
                                                      </table>
                                                    </td>
                                                        </tr>

                                                        <tr>
                                                    <td class="MiddlePack" colspan="2"></td>
                                                        </tr>
													<tr>
														<td class="FieldCellSub" width="133">Reference No</td>
														<td class="FieldCell" width="591">
														<% Response.Write sCreatedMiscPymtNo 
															IF CStr(sCrDrIndi) = "D" Then
																if trim(sPaymentAgainst) = "I" then 
														%>
																	<a class="ExcelDisplayLink" href="#" onClick="ViewInvoiceDetailspopup('<%=iInvNo%>','<%=sInvCode%>'); return false;" >
																		<img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Invoice">
																	</a>	
														<%		End if 
															End IF
														%>
																 
														
															<%	IF CStr(sCrDrIndi) = "D" Then
																	if trim(sPaymentAgainst) = "R" then 
															%>
																	<a class="ExcelDisplayLink" href="#" onClick="showReceiptpopup('<%=sGrnCode%>','<%=iRcptNo%>','<%=sGRNAgainstStr%>','<%=sReceiptRouteStr%>','','<%=sReceiptCode%>','<%=sItemType%>'); return false;" >
																		<img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Receipt">
																	</a>	
															<%
																	End if 
																End IF
															%>
														</td>
				                                     </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Pay For</td>
                                                    <td class="FieldCell" width="591">
													<% Response.Write sPayTo %>
															</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Accounting Head</td>
                                                    <td class="FieldCell" width="591">
                                                            <select size="1" name="selAccHead" class="FormElem" onChange="selAccountHead(this)">
															<option value="A">Select Account Head</option>
									<%
										dim iHeadCount
										'iHeadCount=popFrequentHead(sOrgId,"01",sBookCode)

								    %>
															<option value="G">General Ledger</option>
														<%populatePartyType(sOrgId)%>
                                                    </select>
                                                    </td>
                                                    <input type="hidden" name="hHeadCount" value="<%=iHeadCount%>">
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133"></td>
                                                    <td width="596">
                                                            <span class="DataOnly" id="spAccHead"></span>  </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Pay to / Received from</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtPayTo" size="40" class="Formelem" value="<%=sPayTo%>"  > 
                                                    &nbsp; <a href="#" onclick="SelMisParty(); return false;"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a>
                                                    </td>
                                                        </tr>
                                                        <tr>
                                                    <td width="143" valign="top">
                                                      <table border="0" width="100%" cellspacing="1">
                                                        <tr>
                                                          <td width="50%" class="FieldCellSub">Narration</td>
                                                          <td width="50%" class="FieldCellSub"></td>
                                                        </tr>
                                                      </table>
                                                    </td>
                                                    <td class="FieldCell" valign="top" width="591"> 
                                                    <textarea rows="3" name="txtNarration" cols="50" class="FormElem"><%Response.Write(sPurVouNarr)%> </textarea> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Amount</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtAmount" size="15" style="text-align:right" maxlength="13" class="Formelem" onblur="TDSAmount()" value="<%=sVouAmount%>" > </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Select TDS Group</td>
                                                     <td class="FieldCellSub" width="591"> 
                                                    <select size="1" name="SelTDSGrp" class="FormElem" onchange="TDSAmount()">
                                                    <Option Value="0" selected> Select </option>
		                                                  <% Dim sUseable,sGrpID,sTemp,objRs1
		                                                  Set objRs1 =server.CreateObject("ADODB.Recordset")	
																
																sQuery = "Select GroupID,GroupName from ACC_M_TDSGroup where OUDefinitionID = '"& sOrgId &"' and isNull(Useable,'Y') <> 'N' "
																	'Response.Write sQuery 
																	With objRs1 
																		.CursorLocation = 3
																		.CursorType = 3
																		.ActiveConnection = con
																		.Source = sQuery
																		.Open
																	End With
																	Do while Not objRs1.EOF 
																		sGrpId = objRs1(0)
																	Response.Write objRs1(1)& "<BR>"%>
																	<option value="<%=objRs1(0)%>" <%'If trim(sGroupName) = trim(objRs1(0)) then Response.Write "selected" %>> <%=objRs1(1)%> </option>
																	<%objRs1.MoveNext 
																Loop
																objrs1.Close
															%>
                                                    </select>
                                                    &nbsp; % On Amount &nbsp;
                                                    <input type="text" name="txtTdsAmount" Value="" size="15" style="text-align:right" maxlength="13" class="Formelemread" readonly disabled> 
                                                    <a href="#" onclick="TDSCalc(); return false;"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="TDS Group Selection" width="10" height="11"></a>
                                                     <!--input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled-->
                                                    </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Approval</td>
                                                    <td class="FieldCell" width="591"> <input type="radio" value="Y" checked name="optApprove" class="FormElem" onClick="ResetList(this)">
                                                      Yes&nbsp;&nbsp;
                            <input type="radio" value="N" name="optApprove" class="FormElem" onClick="ResetList(this)"> No &nbsp;&nbsp; Approver &nbsp; 
                            <select size="1" name="selUserId" class="FormElem">
											<option value="I">Immediate Approver</option>
											<%=populateEmployeeWithVal(sUserId)%>
											    </select></td>
                                                        </tr>
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td >
<DIV class=frmBody id="Disaddtional" style="height:1; visibility: hidden;">
<div id="DisCCANL" class=frmBody style="height:1; visibility: hidden;">
	<table cellpadding="0" cellspacing="0" >
		<tr>
			<td class=MiddlePack colspan="3"> </td>
		</tr>
		<tr>
			<td class=FieldCell>
				<DIV class=frmBody id="DisCost" style="width:270;height:100;">
					<table border="0" id="tblCost" cellspacing="1" class="ExcelTable">
						<tr>
							<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								<td class="ExcelHeaderCell" align="center" width="150">Cost Center Head</td>
								<td class="ExcelHeaderCell" align="center">Ratio</td>
								<td class="ExcelHeaderCell" align="center">Amount</td>
						 </tr>
					</table>
				</div><!--End of CostCenter Display Division -->
			</td>
			<td class=ClearPixel width="5">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">                   </td>
			<td class=FieldCell>
				<DIV class=frmBody id="DisAnal" style="width:280; height:100;">

					<table border="0" id="tblAnal" cellspacing="1" class="ExcelTable">
						<tr>
								<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								<td class="ExcelHeaderCell" align="center" width="150">Analytical Head</td>
								<td class="ExcelHeaderCell" align="center">Ratio</td>
								<td class="ExcelHeaderCell" align="center">Amount</td>
					    </tr>
					</table>
				</div>	<!--End of Analytical Display Division -->
			</td>
		</tr>
		<tr>
			<td class=MiddlePack  colspan="3"></td>
		</tr>
	</table>
</div> <!--End of CCANAL Display Division -->
	<DIV class=frmBody id="DisPayable" style="width: 555; visibility: hidden; height:1;">
		<table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="555">
			<tr>
				<td class="ExcelHeaderCell" align="center" rowspan="2" width="10">S.No.</td>
				<td class="ExcelHeaderCell" align="center" colspan="2">Document</td>
				<td class="ExcelHeaderCell" align="center" width="275" colspan="4">Amount</td>
		    </tr>
		   <tr>
				<td class="ExcelHeaderCell" align="center">Detail</td>
				<td class="ExcelHeaderCell" align="center">Date</td>
				<td class="ExcelHeaderCell" align="center">Amount</td>
				<td class="ExcelHeaderCell" align="center">Adjusted</td>
				<td class="ExcelHeaderCell" align="center">To Account</td>
				<td class="ExcelHeaderCell" align="center">To adjust</td>
		   </tr>
		</table>
	</div>
</div><!--End of Addtional Details Display  -->
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
                                                                <input type="Button" value="Add Entry" name="B11" onClick="AddEntry('A')" class="ActionButton">
                                                                 <input type="button" value="Next" name="B12" onClick="AddEntry('S')" class="ActionButton" >
                                                                 <input type="button" value="Cancel" name="btnCancel" onClick="Cancel('MSIVOUBOOKSELECTION.ASP')" class="ActionButton" >
                                                                <input type="reset" value="Reset" name="B14" class="ActionButton" >
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
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
<DIV class=frmBody id="DisVoucher" style="width:100%; visibility:hidden; height:1;">
	<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="100%">
	<tr>
		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		<td class="ExcelHeaderCell" align="center" width="75">AU</td>
		<td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		<td class="ExcelHeaderCell" align="center" width="125">Narration</td>
		<td class="ExcelHeaderCell" align="center" width="125">Amount</td>
		<td class="ExcelHeaderCell" align="center">Additional Details</td>
		<td class="ExcelHeaderCell" align="center">Deduction Amount</td>
		<td class="ExcelHeaderCell" align="center">Deduction Percentage</td>
	</tr>
	</table>
</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
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
</HTML>
<%

Function getReceiptType(sRcptAgainst)
Select Case sRcptAgainst
		Case "01" :	getReceiptType = "Purchase Order"
		Case "02" :	getReceiptType = "Supplier Replacements"
		Case "03" :	getReceiptType ="Supplier Samples For Approval"
		Case "04" :	getReceiptType = "Job Order"
		Case "05" :	getReceiptType = "Job Order Rework/Replacement"
		Case "06" : getReceiptType = "Customer Samples"
		Case "07" :	getReceiptType = "Sales Returns"
		Case "08" : getReceiptType = "Return Of Transferred Goods"
		Case "09" :	getReceiptType = "Inter-unit Transfer"
		Case "10" :	getReceiptType = "Without Reference"
	End Select
End Function

Function getReceiptRoute(sRcptRoute)
Select Case sRcptRoute
	Case "DU" : getReceiptRoute = "Direct User"
	Case "IN" : getReceiptRoute = "Inspection"
	Case "ST" : getReceiptRoute = "Stock"
	Case "ID" : getReceiptRoute = "Inspection-Direct"
	Case "IS" : getReceiptRoute = "Inspection-Stock"
	Case "SD" : getReceiptRoute = "Inspection-Stock-Direct"
End Select
End Function

%>
