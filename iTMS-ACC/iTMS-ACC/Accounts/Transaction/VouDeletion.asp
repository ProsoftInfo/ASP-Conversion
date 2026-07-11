<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDeletion.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	January  24, 2003
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

<!--#include file="../../include/sessionVerify.asp"-->
<%
dim sQuery,iTransNo,objRs,sVouName,iAmdendmentNo,iRecPayNo
Dim Root,ObjDom,sExp,TempNode,iCtr,iPayNo,iCrAdvNo,dAdjAmt

set objRs  = server.CreateObject("adodb.recordset")
Set ObjDom = Server.CreateObject("Microsoft.XMLDOM")

iTransNo=Request.Form("hTransNo")
sVouName=Request.Form("hVouName")

ObjDom.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
SET Root=ObjDom.documentElement




con.BeginTrans

'--------------Insert into Histroy tables---------------------------
sQuery="select isnull(max(AmendmentNo),0)+1 from Acc_T_HistoryVoucherHeader where "&_
		"TransactionNumber="&iTransNo

objRs.open sQuery,con
	iAmdendmentNo=objRs(0)
objRs.close()		

sQuery="INSERT INTO Acc_T_HistoryVoucherHeader(TransactionNumber, AmendmentNo,OUDefinitionID, "&_
	"BookCode, BookNumber, TransactionType, AccountHead, PartyType, "&_
	"PartySubType, PartyCode, VoucherNumber, CreatedVoucherNo, CreatedTransNo, ReceiptNo,"&_
	"VoucherDate, VoucherAmount, CrDrIndication, CreatedBy, CreatedOn, ApprovedBy, ApprovedOn,"&_
	"AccountedBy, AccountedOn, AuditedBy, AuditedOn, BRSTransactionNo, BRSAmendmentNo, ClearedOn,"&_
	"BankInstrumentType, BankInstrumentNo, BankInstrumentDate, DrawnOnBank, PayableAt, "&_
	"HistoryType, HistoryBy, HistoryOn, HistoryReason)"&_
	
	"SELECT CreatedTransNo, "&iAmdendmentNo&", OUDefinitionID, "&_
	"BookCode, BookNumber, TransactionType, AccountHead, PartyType, "&_
	"PartySubType, PartyCode, CreatedVoucherNo,CreatedVoucherNo,CreatedTransNo, ReceiptNo, "&_
	"VoucherDate, VoucherAmount, CrDrIndication, CreatedBy, CreatedOn, ApprovedBy, ApprovedOn,"&_
	"NULL, NULL, NULL, NULL, NULL, NULL, NULL,"&_
	"NULL, NULL, NULL, NULL, NULL, "&_
	"'D',"&getUserid&",getdate(),'Voucher Deleted' "&_
	"FROM Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo	
	
con.execute(sQuery)

sQuery="INSERT INTO Acc_T_HistoryVoucherDetails(TransactionNumber, AmendmentNo, VoucherEntryNumber,AccountingUnit, AccUnitAccountHead, "&_
		"AccUnitPartyType, AccUnitPartySubType,AccUnitPartyCode, VoucherNarration, Amount, TransCrDrIndication) "&_
		"SELECT CreatedTransNo, "&iAmdendmentNo&",VoucherEntryNumber, AccountingUnit, AccUnitAccountHead, AccUnitPartyType, AccUnitPartySubType,"&_
		"AccUnitPartyCode, VoucherNarration, Amount, TransCrDrIndication FROM Acc_T_CreatedVoucherDetails	where "&_
		" CreatedTransNo="&iTransNo			
con.execute(sQuery)	

sQuery="INSERT INTO Acc_T_HistoryVoucherCC(TransactionNumber, AmendmentNo, VoucherEntryNumber, "&_
		"AccountingUnit, AccUnitAccountHead, AccUnitCCHead, CCRatioPercent, CCRatioAmount) "&_
		"select CreatedTransNo,"&iAmdendmentNo&", VoucherEntryNumber, AccountingUnit, "&_
		"AccUnitAccountHead, AccUnitCCHead, CCRatioPercent, CCRatioAmount from Acc_T_CreatedVoucherCCDet "&_
		"where CreatedTransNo="&iTransNo
		
con.execute(sQuery)

sQuery="INSERT INTO Acc_T_HistoryVoucherAH(TransactionNumber, AmendmentNo, VoucherEntryNumber, "&_
		"AccountingUnit, AccUnitAccountHead, AccUnitAnalyticalCode, RatioPercentage, RatioAmount) "&_
		"select CreatedTransNo,"&iAmdendmentNo&", VoucherEntryNumber, AccountingUnit, "&_
		"AccUnitAccountHead, AccUnitAnalyticalCode, RatioPercentage, RatioAmount from Acc_T_CretedVoucherAHDet "&_
		"where CreatedTransNo="&iTransNo
con.execute(sQuery)
		
sQuery="INSERT INTO Acc_T_HistoryPybleAdjustment(PayablesNumber, PaidByTransactionNo, AmendmentNo, "&_
		"PaidOn, AmountPaid)SELECT PayablesNumber, CreatedTransNo,"&iAmdendmentNo&", PaidOn, AmountPaid FROM Acc_T_CreatedPybleAdjDet "&_		
		"where CreatedTransNo="&iTransNo
con.execute(sQuery)		

sQuery="INSERT INTO Acc_T_HistoryRcvblAdjustment(ReceivableNumber, RecdByTransactionNo, AmendmentNo, "&_
		"ReceivedOn, AmountReceived)SELECT ReceivableNumber, CreatedTransNo,"&iAmdendmentNo&", ReceivedOn, AmountReceived "&_
		"FROM Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="INSERT INTO Acc_T_HistoryVoucherAdvPymt(TransactionNumber, AmendmentNo, OUDefinitionID, "&_
	"PartyType, PartySubType, PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, "&_
	"AdvanceAdjusted) SELECT CreatedTransNo,"&iAmdendmentNo&", OUDefinitionID, "&_
	"PartyType, PartySubType, PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived,"&_
	"AdvanceAdjusted FROM Acc_T_CreatedAdvances where CreatedTransNo="&iTransNo
con.execute(sQuery)
'----------Delete Reocords-----------------------------------------
				
sQuery="delete Acc_T_VouchersForApproval where CreatedTransNo="&iTransNo
con.execute(sQuery)
				
sQuery="delete Acc_T_CretedVoucherAHDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherCCDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo

Response.Write sQuery
con.execute(sQuery)

sQuery="delete Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedAdvances where CreatedTransNo="&iTransNo
con.execute(sQuery)

sExp = "//PayRec/Doc[@AdjType=""R""]"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		iPayNo = TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		dAdjAmt = TempNode.Item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value
		
		sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iPayNo&" "
		objRs.Open sQuery,Con
		IF Not objRs.EOF Then
			iCrAdvNo = objRs(0)
		Else
			iCrAdvNo = 0
		End IF
		objRs.Close
		
		sQuery = "UPDATE Acc_T_CreatedAdvances SET AdvanceAdjusted =  "&_
				 "isNull(AdvanceAdjusted,0) - "&dAdjAmt&" WHERE CreatedAdvanceNo = "&iCrAdvNo&" "
				 
		Response.Write sQuery &"<br><br>"
		Con.Execute sQuery		
				
	Next	
End IF

sExp = "//PayRec/Doc[@AdjType=""P""]"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		iPayNo = TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		dAdjAmt = TempNode.Item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value
		
		sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iPayNo&" "
		objRs.Open sQuery,Con
		IF Not objRs.EOF Then
			iCrAdvNo = objRs(0)
		Else
			iCrAdvNo = 0
		End IF
		objRs.Close
		
		sQuery = "UPDATE Acc_T_CreatedAdvances SET AdvanceAdjusted =  "&_
				 "isNull(AdvanceAdjusted,0) - "&dAdjAmt&" WHERE CreatedAdvanceNo = "&iCrAdvNo&" "
		
		Response.Write sQuery &"<br><br>"
		Con.Execute sQuery		
				
	Next	
End IF

sQuery="delete Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
con.execute(sQuery)

'Con.RollbackTrans
'	Response.End

Response.Clear

if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	con.CommitTrans
	'Con.RollbackTrans
	'Response.End
	%>

<SCRIPT>
<!--
	function msgbox(strr)
	 {
			alert(strr);
			window.location.href = document.formname.hTargetPage.value;
     }
//-->
</SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<BODY BGCOLOR="#336699" onLoad = "msgbox('Voucher Deleted Successfully')">
<form id="formname" name="formname">
<input type="hidden" name="hTargetPage" value="<%	
	select case sVouName
	Case "CA" Response.Write "VouCABookSelection.asp"
	Case "BA" Response.Write "VouBABookSelection.asp"
	Case "GJ" Response.Write "VouGJBookSelection.asp"
	end select
end if
%>">
</form>
</BODY>
</HTML>

