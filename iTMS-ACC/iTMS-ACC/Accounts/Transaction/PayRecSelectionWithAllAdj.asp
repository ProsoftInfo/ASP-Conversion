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
	'Program Name				:	PayRecSelectionWithAllAdj.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<%
dim sOrgId,sPartyValue,objRs,sQuery,sVouType,sParCode,sParSubType,sParType,sVouDate
Dim sQuery2,sQuery3,sQuery4,objRs2

sOrgId=Request("orgid")
sVouDate=Request("VouDate")
sPartyValue=split(trim(Request("ParCode")),"?")
sParType=sPartyValue(0)
sParSubType=sPartyValue(1)
sParCode=sPartyValue(3)
sVouType=Request("Type")


Response.Write "<font color=red>"
'Response.Write sParType &" " & sParSubType &" " & sParCode &" " & sVouDate &" " & sOrgId &"<br>"

if trim(sVouDate)="" then
	sVouDate=Day(date)&"/"&Month(date)&"/"&Year(date)
end if

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML id="AccHeadData">
<root No="0">
	<PayRec/>
	<RecCount Val="0" />
</root>
</XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<script language="javascript">
window.__itmsPopupCompat = { type: "payRecSelection", formName: "frm1" };
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onunload="return window_onunload()">
<form method="POST" name="frm1" action="">
<input type=hidden name="hVouType" value="<%=sVouType%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		Adjustments
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
												<DIV class=frmBody id=frm1 style="width: 510; height:380;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="480">
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
 <tr><td colspan="8" class="ExcelSerial"><b>Sales Invoices</b></td></tr>

<%
Response.Write "<font color=red>"

'*********************** Adjustment Type Details ******************************************
'		I  -- Sales Invoices
'		PI -- Purchase Invoices
'		D  -- Debit Notes
'		C  -- Credit Notes
'		P  -- Advance Payments
'		R  -- Advance Receipts
'*********************** Adjustment Type Details ******************************************

dim iDocNo,iSno,iPayNo,iRecCount,sCrDrNoDesc
dim sAmtReceived,sAmtReceiavble,sInvDate,sInvNo,sVouNo,dToAccount,dToAdjusted
Dim sCrDrNo,sCrDrType

sQuery = "Select P.CreatedReceivable,P.Narration,Convert(Char,P.PartyInvoiceDate,103),P.AmountReceivable, "&_
		  "P.AmountReceived,C.AmountReceived - P.AmountReceived,P.ReceivableNumber "&_
		  "From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C Where "&_
		  "P.AmountReceivable > P.AmountReceived and P.OUDefinitionID = '"&sOrgId&"' and P.PartyType='"&sParType&"' and P.PartySubType="&sParSubType&" "&_
		  "and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) "&_
		  "and P.TransactionNumber = H.TransactionNumber and C.ReceivableNumber = P.CreatedReceivable  "&_
		  "and C.AmountReceivable > C.AmountReceived and H.TransactionType = 'SJR' "

'Response.write "<textarea>"& sQuery&"</textarea>"
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

iRecCount = objRs.RecordCount

set objRs.ActiveConnection = nothing

iSno=1
If not objRs.EOF then
	Do While Not objRs.EOF
		sInvNo = objRs(1)
		
		IF Trim(sInvNo) = "&sNewNarr&" Then
			sQuery = "SELECT H.VoucherNumber, Convert(Varchar,H.VoucherDate,103) FROM Acc_T_Receivables R  "&_
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

	'	IF CStr(Trim(sInvNo)) = "&sNewNarr&" Then
'
'		End IF


		dToAdjusted = Cdbl(sAmtReceiavble - sAmtReceived - dToAccount)
			IF CDbl(dToAdjusted) <> 0 Then

%>
			<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?I">
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem"></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%>-<%=iDocNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>

				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td>
            </tr>

<%
			iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
	Loop
Else
%>
<tr>
	<td class="ExcelDisplayCell" align="center" colspan="8"><b>No Sales Invoice Found !!</b></td>
</tr>
<%
end if
objRs.Close
%>
<tr>
<td colspan="8" class="ExcelSerial"><b>
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
'						:	D - Credit Note
'						:	C - Debit Note
'===================================================================================

sQuery2 = "Select P.CreatedReceivable,H.VoucherNumber,Convert(Char,P.PartyInvoiceDate,103),P.AmountReceivable, "&_
		  "P.AmountReceived,C.AmountReceived - P.AmountReceived,P.ReceivableNumber, Convert(Char,H.VoucherDate,103) "&_
		  "From Acc_T_Receivables P, Acc_T_VoucherHeader H, Acc_T_CreatedReceivables C Where "&_
		  "P.AmountReceivable > P.AmountReceived and P.OUDefinitionID = '"&sOrgId&"'  "&_
		  "and P.PartyCode = "&sParCode&" and P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) "&_
		  "and P.TransactionNumber = H.TransactionNumber and C.ReceivableNumber = P.CreatedReceivable  "&_
		  "and C.AmountReceivable > C.AmountReceived and H.TransactionType = 'DNR' "
'Response.Write "<textarea>"& sQuery2 &"</textarea>"
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

			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?D">
            <td class="ExcelSerial" align="center"><%=iSno%></td>
            <td class="ExcelInputCell" align="right" width="10">
            <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem"></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sCrDrVouDate%></td>

            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td-->
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td>
       </tr>

<%
		iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
	Loop
Else
%>
<tr>
	<td class="ExcelDisplayCell" align="center" colspan="8"><b>No Debit Notes Found !!</b></td>
</tr>
<%

end if
objRs.Close

sCrDrNoDesc = "Purchase Biils "
%>

<tr><td colspan="8" class="ExcelSerial"><b>Purchase Bills </b></td></tr>
<%
'********************************** Debit Note Over *****************************************
sQuery3 = " Select a.CreatedPayablesNumber,a.Narration,convert(char,a.PartyBillDate,103), "&_
		  " a.AmountPayable,a.AmountPaid,b.AmountPaid-a.AmountPaid,a.PayablesNumber from Acc_T_Payables a, Acc_T_CreatedPayables b "&_
		  " where a.OUDefinitionID='"&sOrgId&"' and a.PartyType='"&sParType&"'"&_
		  " and b.PayablesNumber=a.CreatedPayablesNumber and a.PartySubType="&sParSubType&" and a.PartyCode="&sParCode&" and a.AmountPayable>a.AmountPaid and a.VoucherDate <= convert(datetime,'"&sVouDate&"',103)"

'Response.write "<textarea>"& sQuery3&"</textarea>"
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery3
	.ActiveConnection = con
	.Open
end with

iRecCount = objRs.RecordCount + CDbl(iRecCount)


set objRs.ActiveConnection = nothing


'iSno=1
If not objRs.EOF then
	Do While Not objRs.EOF
		iDocNo = objRs(0)
		sInvNo = objRs(1)
		sInvDate = objRs(2)
		iPayNo= objRs(6)
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

			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?PI">
            <td class="ExcelSerial" align="center"><%=iSno%></td>
            <td class="ExcelInputCell" align="right" width="10">
            <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem"></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>

            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td-->
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td>
       </tr>

<%
			iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
	Loop
Else
%>
<tr>
	<td class="ExcelDisplayCell" align="center" colspan="8"><b>No Purchase Invoices Found !!</b></td>
</tr>
<%

end if
objRs.Close

sCrDrNoDesc = "Credit Notes "
%>
<tr><td colspan="8" class="ExcelSerial"><b>Credit Notes </b></td></tr>
<%

'**************************** Purchase Invoice Over **************************************
sQuery4 = "Select P.CrCreatedPayable,H.VoucherNumber,Convert(Char,P.PartyBillDate,103), "&_
		  "P.AmountPayable,P.AmountPaid,C.AmountPaid - P.AmountPaid ,P.PayablesNumber, Convert(Char,H.VoucherDate,103) "&_
		  "From Acc_T_Payables P, Acc_T_VoucherHeader H, Acc_T_CreatedPayables C Where "&_
		  "P.CreatedPayablesNumber = 0 and P.Narration is Null and P.AmountPayable > P.AmountPaid "&_
		  "and P.OUDefinitionID = '"&sOrgId&"' and P.PartyCode = "&sParCode&" and "&_
		  "P.VoucherDate <= convert(datetime,'"&sVouDate&"',103) and P.TransactionNumber = "&_
		  "H.TransactionNumber and C.PayablesNumber = P.CrCreatedPayable "&_
		  "And C.AmountPayable > C.AmountPaid "
		  'P.PartySubType = "&sParSubType&" and  "

'Response.Write "<textarea>"& sQuery4 &"</textarea>"

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

			<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?C">
            <td class="ExcelSerial" align="center"><%=iSno%></td>
            <td class="ExcelInputCell" align="right" width="10">
            <input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem"></td>
            <td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
            <td class="ExcelDisplayCell"><p align="center"><%=sCrDrVouDate%></td>

            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
            <!--td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td-->
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td>
       </tr>

<%
			iSno=cint(iSno)+1
		End IF
		objRs.MoveNext
	Loop
Else
%>
<tr>
	<td class="ExcelDisplayCell" align="center" colspan="8"><b>No Credit Notes Found !!</b></td>
</tr>
<%

end if
objRs.Close



'************************ Credit Note Over ************************************************

'************************ Advance Payments/Receipts ****************************************
'IF CStr(sParType) = "CR" Then
sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(char,b.VoucherDate,103),isnull(a.AdvanceAdjusted,0), "&_
		 "isnull(a.AdvancePaid,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
		 "b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
		 "A.AdvanceNumber,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
		 "from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
		 "a.OUDefinitionID='"&sOrgId&"' and a.PartyType='"&sParType&"' and a.PartySubType="&sParSubType&" and a.PartyCode="&sParCode&" and "&_
		 "isnull(a.AdvancePaid,0)>isnull(a.AdvanceAdjusted,0) and "&_
		 "b.TransactionNumber = a.TransactionNumber "&_
		 "and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
		 "A.CreatedAdvanceNo = C.CreatedAdvanceNo "
'Else
'sQuery = "select a.TransactionNumber,b.VoucherNumber,convert(char,b.VoucherDate,103),isnull(a.AdvanceAdjusted,0), "&_
'		 "isnull(a.AdvanceReceived,0),b.TransactionType,b.PayToRecdFrom,b.BankInstrumentType, "&_
'		 "b.BankInstrumentNo,convert(char,b.BankInstrumentDate,103),DrawnOnBank, A.CreatedTransNo, "&_
'		 "A.AdvanceNumber,isNull(C.AdvanceAdjusted,0) - isnull(a.AdvanceAdjusted,0) ToAccount "&_
'		 "from Acc_T_AdvancePayments a, Acc_T_VoucherHeader b, Acc_T_CreatedAdvances C where  "&_
'		 "a.OUDefinitionID='"&sOrgId&"' and a.PartyType='"&sParType&"' and a.PartySubType="&sParSubType&" and a.PartyCode="&sParCode&" and "&_
'		 "isnull(a.AdvanceReceived,0)>isnull(a.AdvanceAdjusted,0) and "&_
'		 "b.TransactionNumber = a.TransactionNumber "&_
'		 "and b.VoucherDate<= convert(datetime,'"&sVouDate&"',103) and "&_
'		 "A.CreatedAdvanceNo = C.CreatedAdvanceNo "
'End IF

'Response.Write sQuery &"<BR>"

'IF CStr(sParType) = "CR" Then
	sCrDrType = "P"
	sCrDrNoDesc = "Advance Payments  "
%>
	<tr>
		<td class="ExcelHeaderCell" align="left" colspan="8"><b>Advance Payments</b></td>
	</tr>
<%'Else
'	sCrDrType = "R"
'	sCrDrNoDesc = "Advance Receipts  "
%>
	<!--tr>
		<td class="ExcelHeaderCell" align="left" colspan="8"><b>Advance Receipts</b></td>
	</tr-->
<%
'End IF

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
'iSno=1

IF Not objRs.EOF Then
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

%>
			<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>">
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem"></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>

				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td>
            </tr>

<%				iSno=cint(iSno)+1
			End IF
			objRs.MoveNext
		Loop
	Else
%>
		<tr>
			<td class="ExcelDisplayCell" align="center" colspan="8"><b>No Advance Payments Found !!</b></td>
		</tr>
<%
	end if
	objRs.Close


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


'Response.Write sQuery &"<BR>"

'IF CStr(sParType) = "CR" Then
'	sCrDrType = "P"
'	sCrDrNoDesc = "Advance Payments  "
%>
	<tr>
		<td class="ExcelHeaderCell" align="left" colspan="8"><b>Advance Receipts</b></td>
	</tr>
<%'Else
	sCrDrType = "R"
	sCrDrNoDesc = "Advance Receipts  "
%>
	<!--tr>
		<td class="ExcelHeaderCell" align="left" colspan="8"><b>Advance Receipts</b></td>
	</tr-->
<%
'End IF

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
'iSno=1

IF Not objRs.EOF Then
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

%>
			<tr>
				<input type="hidden" name="hDoc<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" value="<%=sCrDrNoDesc%><%=sInvNo%>?<%=sInvDate%>?<%=sAmtReceiavble%>?<%=sAmtReceived%>?<%=dToAccount%>?<%=iPayNo%>?<%=sCrDrType%>">
				<td class="ExcelSerial" align="center"><%=iSno%></td>
				<td class="ExcelInputCell" align="right" width="10">
				<input type="checkbox" name="chkDocument" value="<%=iDocNo%>Z<%=iPayNo%>Z<%=iSno%>" class="FormElem"></td>
				<td class="ExcelDisplayCell" align="left"><%=sInvNo%></td>
				<td class="ExcelDisplayCell"><p align="center"><%=sInvDate%></td>

				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceiavble,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(sAmtReceived,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAccount,2,,,0)%></td>
				<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dToAdjusted,2,,,0)%></td>
            </tr>

<%				iSno=cint(iSno)+1
			End IF
			objRs.MoveNext
		Loop
	Else
%>
		<tr>
			<td class="ExcelDisplayCell" align="center" colspan="8"><b>No Advance Receipts Found !!</b></td>
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


