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
	'Program Name				:	PayRecSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	January 03,2002
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
dim sOrgId,sPartyValue,objRs,sQuery,sVouType,sParCode,sParSubType,sParType,sVouDate
Dim sQuery2,sQuery3,sQuery4,sCurrDay,sCurrMon

'Response.Write Request("ParCode")&"="

sOrgId=Request("orgid")
sVouDate=Request("VouDate")
sPartyValue=split(trim(Request("ParCode")),"?")
sParType=sPartyValue(0)
sParSubType=sPartyValue(1)
sParCode=sPartyValue(3)
sVouType=Request("Type")

if trim(sVouDate)="" then
	sCurrDay = Day(date)
	sCurrMon = Month(date)
	IF Len(sCurrDay) = 1 Then
		sCurrDay = "0"&sCurrDay
	End IF

	IF Len(sCurrMon) = 1 Then
		sCurrMon = "0"&sCurrMon
	End IF

	sVouDate=sCurrDay&"/"&sCurrMon&"/"&Year(date)
end if

Set objRs = Server.CreateObject("ADODB.RecordSet")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<root No="0">
	<PayRec/>
	<RecCount Val="0" />
</root>
</script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script>
window.__itmsPopupCompat = { type: "payRecSelection", formName: "frm1" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onunload="return window_onunload()">
<form method="POST" name="frm1" action="">
<input type=hidden name="hVouType" value="<%=sVouType%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%if sVouType="D" then%>
		Receivables
		<%else%>
		Payables
		<%end if%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
			<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
												<DIV class=frmBody id=frm1 style="width: 485; height:380;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="475">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center"  rowspan="2"></td>
                                        <td class="ExcelHeaderCell" align="center" colspan="3">Document</td>
                                        <td class="ExcelHeaderCell" width="100" align="center" colspan="3">Amount </td>

                                            </tr>
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center">Detail</td>
                                        <td class="ExcelHeaderCell" align="center">Date</td>
                                        <td class="ExcelHeaderCell" align="center">Amount</td>
                                        <td class="ExcelHeaderCell" align="center"> Adjusted</td>
                                        <% 'IF CStr(sVouType) = "D" Then %>
                                        <!--td class="ExcelHeaderCell" align="center"> Cr Note Value</td>
                                        <%'Else%>
                                        <td class="ExcelHeaderCell" align="center"> Dr Note Value</td-->
                                        <%'End IF %>

                                        <td class="ExcelHeaderCell" align="center">To Account</td>
                                        <td class="ExcelHeaderCell" align="center">To Adjust</td>

                                            </tr>
<%
dim iDocNo,iSno,iPayNo,iRecCount,sCrDrNoDesc
dim sAmtReceived,sAmtReceiavble,sInvDate,sInvNo,sVouNo,dToAccount,dToAdjusted

if sVouType="D" then
sQuery ="select a.CreatedReceivable,a.Narration,convert(char,a.PartyInvoiceDate,103), "&_
	"a.AmountReceivable,a.AmountReceived,b.AmountReceived-a.AmountReceived,a.ReceivableNumber from Acc_T_Receivables a, Acc_T_CreatedReceivables b "&_
	" where a.OUDefinitionID='"&sOrgId&"' and a.PartyType='"&sParType&"'"&_
	" and b.ReceivableNumber=a.CreatedReceivable and a.PartySubType="&sParSubType&" "&_
	" and a.PartyCode="&sParCode&" and a.AmountReceivable>a.AmountReceived and a.VoucherDate <= convert(datetime,'"&sVouDate&"',103)"

'===================================================================================
'Added On		:	18/11/2004
'Reason			:	To list out all related Debit notes related to party

'===================================================================================
sQuery2 = "Select P.ReceivableNumber,H.VoucherNumber,Convert(Char,P.PartyInvoiceDate,103),P.AmountReceivable, "&_
		 "P.AmountReceived,C.AmountReceived - P.AmountReceived,P.ReceivableNumber, Convert(Char,H.VoucherDate,103) "&_
		 "From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C Where "&_
		 "P.CreatedReceivable = 0 and P.Narration is Null and  "&_
		 "P.AmountReceivable > P.AmountReceived and P.OUDefinitionID = '"&sOrgId&"'  "&_
		 "and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) "&_
		 "and P.TransactionNumber = H.TransactionNumber and C.ReceivableNumber = P.DRCreatedReceivable "
		 'and P.PartySubType = "&sParSubType&" "



else
sQuery ="select a.CreatedPayablesNumber,a.Narration,convert(char,a.PartyBillDate,103), "&_
	"a.AmountPayable,a.AmountPaid,b.AmountPaid-a.AmountPaid,a.PayablesNumber from Acc_T_Payables a, Acc_T_CreatedPayables b "&_
	" where a.OUDefinitionID='"&sOrgId&"' and a.PartyType='"&sParType&"'"&_
	" and b.PayablesNumber=a.CreatedPayablesNumber and a.PartySubType="&sParSubType&" and a.PartyCode="&sParCode&" and a.AmountPayable>a.AmountPaid and a.VoucherDate <= convert(datetime,'"&sVouDate&"',103)"



'===================================================================================
'Added On		:	20/11/2004
'Reason			:	To list out all related Credit notes related to party

'===================================================================================
sQuery2 = "Select P.PayablesNumber,H.VoucherNumber,Convert(Char,P.PartyBillDate,103), "&_
		  "P.AmountPayable,P.AmountPaid,C.AmountPaid - P.AmountPaid ,P.PayablesNumber, Convert(Char,H.VoucherDate,103) "&_
		  "From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "&_
		  "P.CreatedPayablesNumber = 0 and P.Narration is Null and P.AmountPayable > P.AmountPaid "&_
		  "and P.OUDefinitionID = '"&sOrgId&"' and P.PartyCode = "&sParCode&" and "&_
		  "P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) and P.TransactionNumber = "&_
		  "H.TransactionNumber and C.PayablesNumber = P.CrCreatedPayable "
			  'P.PartySubType = "&sParSubType&" and  "
end if



'Response.Write sQuery2
Dim sCrDrNo

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

iRecCount = objRs.RecordCount

set objRs.ActiveConnection = nothing

set iDocNo = objRs(0)
set sInvNo = objRs(1)
set sInvDate = objRs(2)
set iPayNo= objRs(6)
iSno=1
If not objRs.EOF then
	Do While Not objRs.EOF
		sAmtReceiavble = Trim(objRs(3))
		sAmtReceived = Trim(objRs(4))
		dToAccount= Trim(objRs(5))
		'dToAdjusted = Trim(objRs(6))


		sAmtReceiavble = CDbl(sAmtReceiavble)
		sAmtReceived = CDbl(sAmtReceived)
		dToAccount= CDbl(dToAccount)



		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)

			IF CDbl(dToAdjusted) <> 0 Then

%>
			                                          <tr>
							<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>" value="<%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?I">
                                        <td class="ExcelSerial" align="center"><%=iSno%></td>
                                        <td class="ExcelInputCell" align="right" width="10">
                                        <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>" class="FormElem"></td>
                                        <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
                                        <td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>

                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td>

                                            </tr>

<%			End IF


		objRs.MoveNext
		iSno=cint(iSno)+1
	Loop
end if
objRs.Close
%>
<tr>
<td colspan="8" class="ExcelSerial"><b>
<%
	IF CStr(sVouType) = "D" Then
		Response.Write "Debit Notes"
		sCrDrNoDesc = "DEBIT NOTE NO: "
	Else
		Response.Write "Credit Notes"
		sCrDrNoDesc = "CREDIT NOTE NO: "
	End IF
%>

</b></td>
</tr>
<%


Dim sCrDrVouDate
'===================================================================================
'Added On				:	18/11/2004
'Reason					:	To list out all related Credit/Debit notes related to party
'							I - Invoice
'							D - Credit Note
'							C - Debit Note
'===================================================================================

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

If not objRs.EOF then
	Do While Not objRs.EOF
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

										<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sVouType%>">
                                        <td class="ExcelSerial" align="center"><%=iSno%></td>
                                        <td class="ExcelInputCell" align="right" width="10">
                                        <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>" class="FormElem"></td>
                                        <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
                                        <td class="ExcelDisplayCell"><p align="center"><%=sCrDrVouDate%></td>

                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
                                        <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td-->
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
                                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td>
                                            </tr>

<%
		End IF

		objRs.MoveNext
		iSno=cint(iSno)+1
	Loop
Else
%>
<tr>
<%IF CStr(sVouType) = "D" Then%>
<td class="ExcelDisplayCell" align="center" colspan="8"><b>No Debit Notes Found !!</b></td>
<%Else%>
<td class="ExcelDisplayCell" align="center" colspan="8"><b>No Credit Notes Found !!</b></td>
<%End IF %>
</tr>
<%

end if
objRs.Close
%>

                                                </table>
                                               <Input type="hidden" name="hRecCount" value="<%=iRecCount%>">
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
                                                                <input type="button" value="Done" name="B8" class="ActionButton" onClick="finaldone()" >
																<input type="button" value="Cancel" name="B8" class="ActionButton" onClick="finalcancel()" >
																<input type="reset" value="Reset" name="B9" class="ActionButton">
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
</table>
</form>
</BODY>
</HTML>


