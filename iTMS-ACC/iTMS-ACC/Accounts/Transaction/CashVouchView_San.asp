<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	CashVouchView_San.asp
	'Module Name				:	ACCOUNTS (Reports)
	'Author Name				:
	'Created On					:	August 2, 2006
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
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%

'XML DOM Variables
Dim oDOM,nodHeader,Root,objRs,sQuery,objRsTemp
dim sNarration,sAccount,sAddtional,iSno,sNarr
dim dTotal,sOrgId,sAccHead,sType,sAccNo
dim EntryNode,HeaderNode,dAmount,sPartyName
dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
dim iTransNo,iBkHeadCode,bOtherUnit,iTdsAmount
Dim sAddress1, sAddress2,sCity,sState,sPostcode,sTranIndication,sTranEntryIndication
Dim iCreatedBy,sCreatedOn,sVouStatus,sEmpName,iPartyCtrlAcc,sAdjType
Dim iEntryNo,iHeadOfAcc,sHeadOfAccName,iHeadOfAccAmt,iPartyCode,iEntryAmt,iCtr
Dim iNetAmtPaid,iTotRecovered,sNarrFlag,sAddnFlag,iTotAddnlAmt,sAdjFlag,sAdjOn,iTotAdj
Dim sBillType,iVouAmt
iNetAmtPaid = 0
iTotRecovered = 0
sNarrFlag = False
sAddnFlag = False
sAdjFlag = False
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iTransNo=Request("TransNo")
'Response.Write iTransNo

'oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")

'oDOM.Load server.MapPath(GetVouchXML(iTransNo))

'set Root=oDOM.documentElement
'sOrgId=Root.Attributes.Item(0).nodeValue
'sOrgName=Root.Attributes.Item(1).nodeValue
'iBookCode =Root.Attributes.Item(2).nodeValue
'sBookName =Root.Attributes.Item(3).nodeValue
'sVouType=Root.Attributes.Item(4).nodeValue
'sVoucDate=Root.Attributes.Item(5).nodeValue
'iBkHeadCode=Root.Attributes.Item(6).nodeValue

'iVouNo=Root.Attributes.Item(9).nodeValue
'sApprove=Root.Attributes.Item(7).nodeValue

'sOrgId = "010101"




set objRs = Server.CreateObject("ADODB.Recordset")
set objRsTemp = Server.CreateObject("ADODB.Recordset")



'sQuery="select OtherUnitTransaction from vwOrgBookNames where OUDefinitionID = '" & sOrgId & "'"&_
'"and BookNumber="&iBookCode&" and BookCode='01'"

'with objRs
'	.CursorLocation = 3
'	.CursorType = 3
'	.Source = sQuery
'	.ActiveConnection = con
'	.Open
'end with

'if not 	objRs.EOF then
'	bOtherUnit=objRs(0)
'else
'	bOtherUnit=0
'end if
'objRs.Close

'==============================Voucher View header queries starts here ============================

sQuery = "select Distinct CreatedVoucherNo,TransactionType,OUDefinitionID,PayToRecdFrom," &_
		"PartyCode,CreatedBy,Convert(varchar,CreatedOn,103),CreatedVouchStatus" &_
		" from VW_Created_CashVoucherView where CreatedTransNo="& iTransNo

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	iVouNo = objRs(0)
	sVouType = objRs(1)
	sOrgId = objRs(2)
	sPayTo = objRs(3)
	iPartyCode = objRs(4)
	iCreatedBy = objRs(5)
	sCreatedOn = objRs(6)
	sVouStatus = objRs(7)
end if
objRs.Close

sQuery = "Select isnull(PurchaseBillType,''),VoucherAmount from Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
'Response.Write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
if not 	objRs.EOF then
	sBillType = objRs(0)
	iVouAmt	  = objRs(1)
End If
objRs.Close
'Response.Write "sBillType="&sBillType
If trim(sVouStatus) = "010104" Then
	sQuery = "select VoucherNumber from ACC_T_VoucherHeader where CreatedTransNo="&iTransNo
	'Response.Write sQuery
	with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
	end with
	if not 	objRs.EOF then
		iVouNo = objRs(0)
	End If
	objRs.Close
End If
set objRs.ActiveConnection = nothing
If trim(sVouType) = "CAP" Then
	sTranIndication = "C"
	sTranEntryIndication = "D"
Else
	sTranIndication = "D"
	sTranEntryIndication = "C"
End If
sQuery = "Select VoucherEntryNumber,AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"' order by 1"
' Response.Write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
'Response.write sQuery
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	iEntryNo = objRs(0)
	iHeadOfAcc = objRs(1)
	iHeadOfAccAmt = FormatNumber(objRs(2),2,,,-2)
	sNarr = objRs(3)
	iPartyCtrlAcc = objRs(4)
end if
objRs.Close

'Response.Write sNarr

IF iHeadOfAcc <> "" Then
	sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
Else
	sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
End If
'Response.write sQuery
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	sHeadOfAccName = objRs(0)
end if
objRs.Close

sQuery = "select OrgUnitDescription,Address1,Address2,City,State,PostCode from DCS_OrganizationUnitDefinitions where OUDefinitionID='"&sOrgId&"'"
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
if not 	objRs.EOF then
	sOrgName = objRs(0)
	sAddress1 = objRs(1)
	sAddress2 = objRs(2)
	sCity = objRs(3)
	sState = objRs(4)
	sPostcode = objRs(5)
else
	sOrgName = ""
	sAddress1 = ""
	sAddress2 = ""
	sCity = ""
	sState = ""
	sPostcode = ""
end if
objRs.Close

If iPartyCode <> "0" Then
	sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCode

	With Objrs
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.Open
	End With
	set objRs.ActiveConnection = nothing
	IF not objRs.EOF  Then
		sPartyName = Trim(Objrs(0))
	End IF
	objRs.Close
End If

	sQuery = "SELECT LoginId FROM DCS_User WHERE InternalUserID ="&iCreatedBy

With Objrs
	.ActiveConnection = con
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.Open
End With
set objRs.ActiveConnection = nothing
IF not objRs.EOF  Then
	sEmpName = Trim(Objrs(0))
End IF
objRs.Close

'==============================Voucher View header queries ends here ============================

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script>
function cashVoucherField(name)
{
	var frm = document.formname;
	var wanted;
	var i;
	if (!frm || !frm.elements) {
		return null;
	}
	if (frm.elements[name]) {
		return frm.elements[name];
	}
	wanted = String(name).toLowerCase();
	for (i = 0; i < frm.elements.length; i += 1) {
		if (String(frm.elements[i].name || "").toLowerCase() === wanted) {
			return frm.elements[i];
		}
	}
	return null;
}

function cashVoucherValue(name)
{
	var item = cashVoucherField(name);
	return item ? item.value : "";
}

function openCashVoucherDialog(url)
{
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, "", "dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No", function () {});
		return;
	}
	window.open(url, "_blank", "width=300,height=200,resizable=no,status=no");
}

function updateCashVoucherApprover()
{
	var approver = cashVoucherField("hApprover");
	var user = cashVoucherField("selUserid");
	var transNo = cashVoucherValue("hTransNo");
	var xhr;
	if (!approver || approver.value !== "Y") {
		return true;
	}
	if (!user || user.selectedIndex <= 0) {
		alert("Select Approver");
		if (user) {
			user.focus();
		}
		return false;
	}
	xhr = new XMLHttpRequest();
	xhr.open("POST", "XMLVouAppUpdate.asp?BkCode=CA&TransNo=" + encodeURIComponent(transNo) + "&User=" + encodeURIComponent(user.value) + "&Mode=E", false);
	xhr.send(null);
	if (String(xhr.responseText || "").replace(/^\s+|\s+$/g, "") !== "") {
		alert(xhr.responseText);
		return false;
	}
	return true;
}

function FinalCheck(flag)
{
	var frm = document.formname;
	var transNo;
	var orgName;
	var vouType;
	var value;
	if (!updateCashVoucherApprover()) {
		return false;
	}
	if (flag === "B") {
		frm.action = "VouCABookSelection.asp";
		frm.submit();
		return false;
	}
	if (flag === "PV" || flag === "RV") {
		frm.action = "VouCAEntry.asp";
		cashVoucherField("selVouType").value = flag === "PV" ? "C" : "D";
		frm.submit();
		return false;
	}
	if (flag === "P") {
		transNo = cashVoucherValue("hTransNo");
		orgName = cashVoucherValue("hOrgName");
		vouType = cashVoucherValue("selVouType");
		value = transNo + ":" + orgName;
		openCashVoucherDialog((vouType === "D" ? "PRNCashRecpVouView.asp?Value=" : "PRNCashPayVouView.asp?Value=") + encodeURIComponent(value));
	}
	return false;
}

function CheckPrint()
{
	openCashVoucherDialog("PRNCashRecVouView2.asp?Value=" + encodeURIComponent(cashVoucherValue("hTransNo")));
}

function CheckPrintStat()
{
	openCashVoucherDialog("PRNCashPayVouView2New.asp?Value=" + encodeURIComponent(cashVoucherValue("hTransNo")));
}
</script>
</HEAD>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action="">
<input type="hidden" name="selUnitId" value="<%=sOrgId%>">
<input type="hidden" name="horgName" value="<%=sOrgName%>">
<input type="hidden" name="hNarr" value="<%=sNarr%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="selBook" value="<%=iBookCode%>">
<input type="hidden" name="selVouType" value="<%=sVouType%>">
<input type="hidden" name="hBookOtherUnit" value="<%=bOtherUnit%>">

<input type="hidden" name="hBookAccHead" value="<%=iBkHeadCode%>">

<input type="hidden" name="hApprover" value="<%=sApprove%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">
					  <%
						if sVouType="CAP" then
							Response.Write "Cash Payment Voucher"
						else
							Response.Write "Cash Receipt Voucher"
						end if
					%>
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">
											<tr>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sOrgName%> </span>
												</td>
												<td class="FieldCellSub" width="80">Voucher No.
												</td>
												<td class="FieldCellSub" width="130">
													<span class="DataOnly"><%=iVouNo %> </span>
												</td>
											</tr>

											<tr>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sAddress1%><br><%=sAddress2%><br><%=sCity%>,<%=sState%>-<%=sPostcode%></span>
												</td>
												<td class="FieldCellSub" valign="top" width="80">Date
												</td>
												<td class="FieldCellSub" valign="top" width="130">
													<span class="DataOnly"><%=sVoucDate%> </span>
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">
											<tr>
												<td class="FieldCellSub" width="100">Head of Account
												</td>
												<td class="FieldCellSub" colspan="3">
													<span class="DataOnly"><%=sHeadOfAccName%> </span>
												</td>
											</tr>

											<tr>
												<td class="FieldCellSub" width="100">Amount
												</td>
												<td class="FieldCellSub" colspan="3">
												<%IF trim(sBillType) = "C" then %>
														<span class="DataOnly"><%=AmountWords(replace(iVouAmt,",",""))%> </span>
												<%Else %>
														<span class="DataOnly"><%=AmountWords(replace(iHeadOfAccAmt,",",""))%> </span>
												<%End IF %>
												</td>
											</tr>

											<tr>
												<td class="FieldCellSub" width="115">
														 <%
																if sVouType="CAP" then
																	Response.Write "Cash Paid To "
																else
																	Response.Write "Cash Received From "
																end if
														%>
												</td>
												<td class="FieldCellSub" colspan=2>
														<%if sType="P" then%>
															<span class="DataOnly"><%=sPartyName %> </span>
														<%else%>
															<span class="DataOnly"><%=sPayTo %> </span>
														<%end if%>
												</td>
												<%IF trim(sBillType) = "C" then %>
													<td class="FieldCellSub" colspan=2><b>Rs.&nbsp;&nbsp;</b><span class="DataOnly"><%=FormatNumber(iVouAmt,2,,,-2)%> </span>
												<%Else %>
													<td class="FieldCellSub" colspan=2><b>Rs.&nbsp;&nbsp;</b><span class="DataOnly"><%=iHeadOfAccAmt%> </span>
												<%End IF %>
												</td>

											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top">
										<table border="0" cellspacing="0" cellpadding="0" width="100%">
											<tr>
												<td class="FieldCell">Details
												</td>
												<td class="FieldCellSub">Received Payment
												</td>
											</tr>
											<tr>
												<td class="FieldCell" valign="top">
													<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">

														<%
															'Fetch the additional Payment / Receipt entries
															sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and VoucherEntryNumber <> "&iEntryNo&" and TransCrDrIndication='"&sTranIndication&"' and Amount <> 0"
															'Response.Write sQuery
															iCtr = 1
															iTotAddnlAmt = 0
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing
															if not 	objRs.EOF then
																Do While not objRs.EOF

																	iHeadOfAcc = objRs(0)
																	iEntryAmt = cdbl(objRs(1))
																	iPartyCtrlAcc = objRs(3)
																	'iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
																	iTotRecovered = iTotRecovered + iEntryAmt
																	sDetFlag = True
																	If iHeadOfAcc <> "" Then
																		sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
																	Else
																		sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
																	End If

																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close

																	If Not sNarrFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2><span class="DataOnly"><b><%=UCase(trim(sNarr))%></b></span></td>
														</tr>

														<%
																	sNarrFlag = True
																	End If
															If Not sAddnFlag Then
																If sVouType = "CAP" Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Recoveries</u></b></span></td>
														</tr>
														<%
																Else
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Additonal Receipts</u></b></span></td>
														</tr>
														<%
																End If
															sAddnFlag = True
															End If
														%>
														<tr>
															<td class="FieldCellSub" colspan="2"><%=sHeadOfAccName%></td>
															<td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
														<%
																iCtr = iCtr + 1

																If iCtr < objrs.RecordCount Then
																	objRs.MoveNext
																	If Not objRs.EOF Then
																		iHeadOfAcc = objRs(0)
																		iEntryAmt = cdbl(objRs(1))
																		'iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
																		iTotRecovered = iTotRecovered + iEntryAmt
																		sDetFlag = True
																		If iHeadOfAcc <> "" Then
																			sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
																		Else
																			sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
																		End If
																		with objRsTemp
																			.CursorLocation = 3
																			.CursorType = 3
																			.Source = sQuery
																			.ActiveConnection = con
																			.Open
																		end with
																		Set objRsTemp.ActiveConnection = Nothing
																		sHeadOfAccName = ""
																		if not 	objRsTemp.EOF then
																			sHeadOfAccName = objRsTemp(0)
																		end if
																		objRsTemp.Close
														%>

															<td class="FieldCellSub" colspan="2"><%=sHeadOfAccName%></td>
															<td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>

														<%
																	End If
																End If
														%>
														</tr>
														<%
																'If iCtr < objrs.RecordCount Then
																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close


														%>


													<%	Dim sDetFlag,sItemDesc
															'Fetch the Recoveries
														'	Response.Write sBillType
														If Trim(sBillType) = "C" then
															sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode,isNull(ItemDescription,'') from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and TransCrDrIndication='"&sTranEntryIndication&"' and Amount <> 0"

															' Response.Write sQuery
															iCtr = 1
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing
															if not 	objRs.EOF then
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	iEntryAmt = cdbl(objRs(1))
																	'iTotRecovered = iTotRecovered + iEntryAmt
																	iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
																	sItemDesc = objRs(4)
																'	Response.Write iHeadOfAcc
																	sDetFlag = True
																	If iHeadOfAcc <> "" Then
																		sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
																	Else
																		sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
																	End If
																	'Response.Write sQuery
																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close

																	If trim(sItemDesc) <> "" then sHeadOfAccName = sItemDesc &"-"&sHeadOfAccName
																If Not sNarrFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2><span class="DataOnly"><b><%=UCase(trim(sNarr))%></b></span></td>
														</tr>

														<%
																	sNarrFlag = True
																End If
																	If iCtr = 1 Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Additional Payments</u></b></span></td>
														</tr>
														<%
																	End If
														%>
														<tr>
															<td class="FieldCellSub" colspan="2"><%=sHeadOfAccName%></td>
															<td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
														</tr>
														<%
																	'If iCtr < objrs.RecordCount Then

																	objRs.MoveNext
																	iCtr = iCtr + 1
																	Loop
																End IF
																objRs.Close
																'To display Tax Entries from Acc_T_CreatedVoucherTaxDet (Tax table)
																sQuery = "select AccountHead,TaxEntryNo from Acc_T_CreatedVoucherTaxDet where createdTransno = "&iTransNo&"  and TaxAmount <> 0  order by 2 "
																objRs.Open sQuery,con
																Do while not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	sDetFlag = True
																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		If iHeadOfAcc <> "" Then
																			.Source = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
																		Else
																			.Source = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
																		End If
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close

																	sQuery = "Select Sum(TaxAmount) from Acc_T_CreatedVoucherTaxDet where CreatedTransNo="&iTransNo&" and  AccountHead = "&iHeadOfAcc&" "
																	objRsTemp.Open sQuery,con
																	If not objRsTemp.EOF then
																		iEntryAmt = cdbl(objRsTemp(0))
																		'iTotRecovered = iTotRecovered + iEntryAmt
																		iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
																	End IF
																	objRsTemp.Close


																If Not sNarrFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2><span class="DataOnly"><b><%=UCase(trim(sNarr))%></b></span></td>
														</tr>

														<%
																	sNarrFlag = True
																End If
																	If iCtr = 1 Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Additional Payments</u></b></span></td>
														</tr>
														<%
																	End If
														%>
														<tr>
															<td class="FieldCellSub" colspan="2"><%=sHeadOfAccName%></td>
															<td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
														</tr>
														<%
																	'If iCtr < objrs.RecordCount Then
																	objRs.MoveNext
																	iCtr = iCtr + 1
																	Loop

																objRs.Close

														Else
															sQuery = "Select AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode from VW_Created_CashVoucherView where CreatedTransNo="&iTransNo&" and VoucherEntryNumber <> "&iEntryNo&" and TransCrDrIndication='"&sTranEntryIndication&"' and Amount <> 0"

															' Response.Write sQuery
															iCtr = 1
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing
															if not 	objRs.EOF then
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	iEntryAmt = cdbl(objRs(1))
																	'iTotRecovered = iTotRecovered + iEntryAmt
																	iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
																'	Response.Write iHeadOfAcc
																	sDetFlag = True
																	If iHeadOfAcc <> "" Then
																		sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
																	Else
																		sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
																	End If
																	'Response.Write sQuery
																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close
																	If Not sNarrFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2><span class="DataOnly"><b><%=UCase(trim(sNarr))%></b></span></td>
														</tr>

														<%
																	sNarrFlag = True
																	End If
																	If iCtr = 1 Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Additional Payments</u></b></span></td>
														</tr>
														<%
																	End If
														%>
														<tr>
															<td class="FieldCellSub" colspan="2"><%=sHeadOfAccName%></td>
															<td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
														<%
																iCtr = iCtr + 1

																If iCtr < objrs.RecordCount Then
																	objRs.MoveNext
																	If Not objRs.EOF Then
																		iHeadOfAcc = objRs(0)
																		iEntryAmt = cdbl(objRs(1))
																		'iTotRecovered = iTotRecovered + iEntryAmt
																		iTotAddnlAmt = iTotAddnlAmt + iEntryAmt
																		sDetFlag = True
																		If iHeadOfAcc <> "" Then
																			sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
																		Else
																			sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
																		End If
																		with objRsTemp
																			.CursorLocation = 3
																			.CursorType = 3
																			.Source = sQuery
																			.ActiveConnection = con
																			.Open
																		end with
																		Set objRsTemp.ActiveConnection = Nothing
																		sHeadOfAccName = ""
																		if not 	objRsTemp.EOF then
																			sHeadOfAccName = objRsTemp(0)
																		end if
																		objRsTemp.Close
														%>

															<td class="FieldCellSub" colspan="2"><%=sHeadOfAccName%></td>
															<td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>

														<%
																	End If
																End If
														%>
														</tr>
														<%
																'If iCtr < objrs.RecordCount Then
																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close
														End If ' If trim(sBillType) = "C" then
															'If sVouType = "CAP" Then
															If trim(sBillType) = "C" then
																iNetAmtPaid =  iTotAddnlAmt
															Else
																iNetAmtPaid = iHeadOfAccAmt + iTotAddnlAmt - iTotRecovered
															End IF
															'Else
															'	iNetAmtPaid = iHeadOfAccAmt - iTotAddnlAmt + iTotRecovered
															'End If

														%>

														<%
															If sNarrFlag Then
														%>
														<tr>
															<td class="FieldCellSub" colspan="2" align="right"><b>Total Recovered Rs.</b></td>
															<td class="FieldCellSub" colspan="2" >
																<span class="DataOnly"><b><%=FormatNumber(iTotRecovered,2,,,-2)%></b></span>
															</td>
														</tr>
														<%End If%>
														<%
															'Fetch the Receivable Adjustments
															sQuery = "select ReceivableNumber,Convert(varchar,ReceivedOn,103),AmountReceived,AdjustType from Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
															'Response.Write sQuery
															iCtr = 1
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing
															if not 	objRs.EOF then
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	sAdjOn = objRs(1)
																	iEntryAmt = cdbl(objRs(2))
																	sAdjType = objRs(3)
																	iTotAdj = iTotAdj + iEntryAmt
																	sDetFlag = True
																	sQuery = "select Narration from ACC_T_CreatedReceivables where Receivablenumber = "&iHeadOfAcc
																	'Response.Write sQuery
																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close
																	'sHeadOfAccName = sHeadOfAccName &" "& sAdjOn
																	If Not sNarrFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2><span class="DataOnly"><b><%=UCase(trim(sNarr))%></b></span></td>
														</tr>

														<%
																	sNarrFlag = True
																	End If
																	If Not sAdjFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Adjustments</u></b></span></td>
														</tr>
														<%
																		sAdjFlag = True
																	End If
															If trim(sAdjType) = "A" Then sHeadOfAccName = "Less Advance Receipts "
														%>
														<tr>
															<td class="FieldCellSub" colspan="2"><%=sHeadOfAccName%>:</td>
															<td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
														</tr>
														<%
																'If iCtr < objrs.RecordCount Then
																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close


														%>
														<%
															'Fetch the Payable Adjustments
															sQuery = "select PayablesNumber,Convert(varchar,PaidOn,103),AmountPaid,AdjustType from Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo
															iCtr = 1
															with objRs
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs.ActiveConnection = nothing
															if not 	objRs.EOF then
																Do While not objRs.EOF
																	iHeadOfAcc = objRs(0)
																	sAdjOn = objRs(1)
																	iEntryAmt = cdbl(objRs(2))
																	sAdjType = objRs(3)
																	iTotAdj = iTotAdj + iEntryAmt
																	sDetFlag = True
																	sQuery = "select Narration from ACC_T_CreatedPayables where Payablesnumber ="&iHeadOfAcc
																	with objRsTemp
																		.CursorLocation = 3
																		.CursorType = 3
																		.Source = sQuery
																		.ActiveConnection = con
																		.Open
																	end with
																	Set objRsTemp.ActiveConnection = Nothing
																	sHeadOfAccName = ""
																	if not 	objRsTemp.EOF then
																		sHeadOfAccName = objRsTemp(0)
																	end if
																	objRsTemp.Close
																	'sHeadOfAccName = sHeadOfAccName &" "& sAdjOn
																	If Not sNarrFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2><span class="DataOnly"><b><%=UCase(trim(sNarr))%></b></span></td>
														</tr>

														<%
																	sNarrFlag = True
																	End If
																	If Not sAdjFlag Then
														%>
														<tr>
															<td class="FieldCell" valign="top" colspan=2>&nbsp;<b><u>Adjustments</u></b></span></td>
														</tr>
														<%
																		sAdjFlag = True
																	End If
															If trim(sAdjType) = "A" Then sHeadOfAccName = "Less Advance Payments "
														%>
														<tr>
															<td class="FieldCellSub" colspan="2"><%=sHeadOfAccName%>:</td>
															<td class="FieldCellSub" colspan="2"><span class="DataOnly"><%=FormatNumber(iEntryAmt,2,,,-2)%></span>
															</td>
														</tr>

														<%
																'If iCtr < objrs.RecordCount Then
																objRs.MoveNext
																iCtr = iCtr + 1
																Loop
															end if
															objRs.Close


														%>
														<%
														IF sNarrFlag <> True then %>
														<tr>
															<td class="FieldCell" valign="top" colspan=2 ><span class="DataOnly"><b><%=UCase(trim(sNarr))%></b></span></td>
														</tr>


														<%End If %>
													</table>
												</td>
												<td class="FieldCellSub" rowspan="2" valign="top">
													<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">
													<%IF sDetFlag <> True then %>
														<tr>
															<td class="FieldCellSub">
															</td>
															<td class="FieldCellSub">
															</td>
															<td class="FieldCellSub">
																<span class="DataOnly"> <%=FormatNumber((iHeadOfAccAmt + iTotAddnlAmt),2,,,-2)%></span>
															</td>
														</tr>

													<%ElseIF sDetFlag = True then
													'Response.Write iHeadOfAccAmt & "---"&iNetAmtPaid&"<BR>"%>
														<tr>
															<td class="FieldCellSub">Total Amount Rs.
															</td>
															<td class="FieldCellSub">
															<%IF trim(sBillType) = "C" then %>
																<span class="DataOnly"><%=FormatNumber((iTotAddnlAmt),2,,,-2)%></span>
															<%Else%>
																<span class="DataOnly"><%=FormatNumber((iHeadOfAccAmt + iTotAddnlAmt),2,,,-2)%></span>
															<%End IF%>
															</td>
														</tr>

														<tr>
															<td class="FieldCellSub">Total Recovered Rs.
															</td>
															<td class="FieldCellSub">
																<span class="DataOnly"><%=FormatNumber(iTotRecovered,2,,,-2)%></span>
															</td>
														</tr>

														<tr>
															<td class="FieldCellSub">Net Paid Rs.
															</td>
															<td class="FieldCellSub">
																<span class="DataOnly"><%=FormatNumber(iNetAmtPaid,2,,,-2)%></span>
															</td>
														</tr>
													<% End If %>
														<tr>
															<td class="FieldCellSub" colspan="2" align="center">
																<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly">
																	<tr>
																		<td class="FieldCellSub">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																			<p>&nbsp;
																			<p>&nbsp;
																		</td>
																	</tr>

																</table>
															</td>
														</tr>

														<tr>
															<td class="FieldCellSub" colspan="2" align="center">Signature
															</td>
														</tr>

													</table>
												</td>
											</tr>

											<tr>
												<td class="FieldCell" valign="bottom">
													<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">
														<tr>
															<td class="FieldCellSub">
																<span class="DataOnly"><%=sEmpName%>-<%=sCreatedOn%></span>
															</td>
															<td class="FieldCellSub">
																<span class="DataOnly">-</span>
															</td>
															<td class="FieldCellSub">
																<span class="DataOnly">-</span>
															</td>
														</tr>

														<tr>
															<td class="FieldCellSub">Prepared By
															</td>
															<td class="FieldCellSub">Checked By
															</td>
															<td class="FieldCellSub">Passed By
															</td>
														</tr>

													</table>
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													    <input type="button" value="Print in Stationary" name="B9" onClick="CheckPrintStat()" class="ActionButtonX">
														<input type="button" value="Print in Plain Paper" name="B8" onClick="CheckPrint()" class="ActionButtonX">
														<input type="button" value="Close" name="B3" class="ActionButton" onclick="window.close()">

												</td>
											</tr>

										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="BottomPack">
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
</body>
