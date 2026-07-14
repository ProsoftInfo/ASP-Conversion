<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PurVouDeletion.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	
	'Created On					:	
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
<%
dim sQuery,iTransNo,objRs,sVouName,iAmdendmentNo,oDOM,sExp,TempNode,iCtr,i,sTransNo
set objRs  = server.CreateObject("adodb.recordset")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")



sTransNo=Request.Form("hTransno")
iTransNo=Split(sTransNo,"|")

'Response.Write sTransNo

for i=0 to UBound(iTransNo)
	Response.Write iTransNo(i) &"<br>"
Next

con.BeginTrans

for i=0 to UBound(iTransNo)
	if CStr(iTransNo(i))<>"0" then
		oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo(i)&".xml")

		sExp = "//AdvanceDetails/Advance[@AdjType=""I""]" 'Only For Cash/Bank Adjustments
		Set TempNode = oDOM.selectNodes(sExp)
		
		'Response.Write iTransNo(i)



		

		'--------------Insert into Histroy tables---------------------------
		sQuery="select isnull(max(AmendmentNo),0)+1 from Acc_T_HistoryVoucherHeader where "&_
				"TransactionNumber="&iTransNo(i)

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
			"FROM Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo(i)	
			
		con.execute(sQuery)

		sQuery="INSERT INTO Acc_T_HistoryVoucherDetails(TransactionNumber, AmendmentNo, VoucherEntryNumber,AccountingUnit, AccUnitAccountHead, "&_
				"AccUnitPartyType, AccUnitPartySubType,AccUnitPartyCode, VoucherNarration, Amount, TransCrDrIndication) "&_
				"SELECT CreatedTransNo, "&iAmdendmentNo&",VoucherEntryNumber, AccountingUnit, AccUnitAccountHead, AccUnitPartyType, AccUnitPartySubType,"&_
				"AccUnitPartyCode, VoucherNarration, Amount, TransCrDrIndication FROM Acc_T_CreatedVoucherDetails	where "&_
				" CreatedTransNo="&iTransNo(i)			
		con.execute(sQuery)	

		sQuery="INSERT INTO Acc_T_HistoryVoucherCC(TransactionNumber, AmendmentNo, VoucherEntryNumber, "&_
				"AccountingUnit, AccUnitAccountHead, AccUnitCCHead, CCRatioPercent, CCRatioAmount) "&_
				"select CreatedTransNo,"&iAmdendmentNo&", VoucherEntryNumber, AccountingUnit, "&_
				"AccUnitAccountHead, AccUnitCCHead, CCRatioPercent, CCRatioAmount from Acc_T_CreatedVoucherCCDet "&_
				"where CreatedTransNo="&iTransNo(i)										
				
		con.execute(sQuery)

		sQuery="INSERT INTO Acc_T_HistoryVoucherAH(TransactionNumber, AmendmentNo, VoucherEntryNumber, "&_
				"AccountingUnit, AccUnitAccountHead, AccUnitAnalyticalCode, RatioPercentage, RatioAmount) "&_
				"select CreatedTransNo,"&iAmdendmentNo&", VoucherEntryNumber, AccountingUnit, "&_
				"AccUnitAccountHead, AccUnitAnalyticalCode, RatioPercentage, RatioAmount from Acc_T_CretedVoucherAHDet "&_
				"where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)
				
		sQuery = "INSERT INTO Acc_T_HisCreatedPayables (AmendmentNo, PayablesNumber, CreatedTransNo,  "&_
				 "OUDefinitionID, PartyType, PartySubType, PartyCode, VoucherDate, PartyBillNumber,  "&_
				 "PartyBillDate, AmountPayable, AmountPaid, Narration) "&_
				 "Select "&iAmdendmentNo&",PayablesNumber, CreatedTransNo, OUDefinitionID, PartyType, PartySubType, PartyCode, "&_
				 "VoucherDate, PartyBillNumber, PartyBillDate,AmountPayable, AmountPaid, Narration "&_
				 "FROM Acc_T_CreatedPayables Where CreatedTransNo = "&iTransNo(i)
				 
		'Response.Write sQuery
				 
		con.execute(sQuery)

		sQuery = "INSERT INTO Acc_T_HistoryCreatedVoucherTaxDet (AmendmentNo, CreatedTransNo, AccountHead, "&_
				 "TaxCode, TaxCategoryCode, TaxEntryNo, InvoiceType, FormNumber, TaxPercentage, TaxAmount, "&_
				 "TransCrDrIndication, TransactTaxAmount) "&_
				 "Select "&iAmdendmentNo&", CreatedTransNo, AccountHead, TaxCode, TaxCategoryCode, TaxEntryNo, InvoiceType, "&_
				 "FormNumber, TaxPercentage, TaxAmount, TransCrDrIndication, TransactTaxAmount "&_
				 "FROM Acc_T_CreatedVoucherTaxDet Where CreatedTransNo = "&iTransNo(i)
				 
		Con.Execute (sQuery)
		'----------Delete Reocords-----------------------------------------
						
		sQuery="delete Acc_T_VouchersForApproval where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)
		
		
						
		sQuery="delete Acc_T_CretedVoucherAHDet where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)

		sQuery="delete Acc_T_CreatedVoucherCCDet where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)

		sQuery="delete Acc_T_CreatedVoucherTaxDet where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)

		sQuery="delete Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)

		sQuery="delete Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)

		sQuery="delete Acc_T_CreatedPayables where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)

		sQuery="delete Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)

		sQuery="delete Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo(i)
		con.execute(sQuery)

		Response.Write sQuery &"<br>"

		IF TempNode.Length <> 0 Then
			For iCtr = 0 To TempNode.Length - 1
				IF CStr(TempNode.Item(iCtr).Attributes.getNamedItem("AmountToAdj").Value) <> "0" Then
					sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = AdvanceAdjusted - "&TempNode.Item(iCtr).Attributes.getNamedItem("AmountToAdj").Value&" Where "&_
						 "CreatedTransNo = "&TempNode.Item(iCtr).Attributes.getNamedItem("CreatedTransNo").Value
						 
					'Response.Write sQuery &"<br>"
					Con.Execute sQuery
				End IF
			Next
		End IF
		
	end if
next

if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	con.CommitTrans
	'con.rollbackTrans
	Response.Redirect("PurchaseVouchers.asp")		
end if
%>


