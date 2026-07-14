<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNSalReturnGenAmd2.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	April 18, 2003
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
<!--#include virtual="/include/NoSeries.asp"-->
<%

'XML DOM Variables
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot,oNodTemp,sQuery
dim sParCode,sParSubType,sParType
dim EntryNode,HeaderNode,nodANL,newElem
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus,iCounter
dim iSeriesNo,iSeriesCode
dim sSalType, ActualTransNo
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
dim dAdjAmtTotal,sAppTy,sAppBy,iCrTransNo,objfs,sCallTy,iAmdendmentNo,iItemCode,iClassCode



sAppTy = Request.Form("optApprove")
sCallTy = Request.Form("hCallType")



sVouCode="07"
sVouType="C"

ActualTransNo = Request("hdTransNo")
sCallTy = Request.Form("hCallType")

set objRs  = server.CreateObject("adodb.recordset")
Set objfs = CreateObject("Scripting.FileSystemObject")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.load server.MapPath("../temp/transaction/Voucher Entry_CNAmd_"&Session.SessionID&".xml")
set oNodRoot=oDOM.documentElement

iCrTransNo = oNodRoot.Attributes.Item(0).nodeValue
iVouNo = oNodRoot.Attributes.Item(1).nodeValue
ActualTransNo = oNodRoot.Attributes.Item(2).nodeValue

for each oNodTemp in oNodRoot.childNodes

	if oNodTemp.nodeName="Header" then
		for Each oNodEntry in  oNodTemp.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Book" then
				sBookNo=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Party" then
				sParType=oNodEntry.Attributes.Item(0).nodeValue
				sParSubType=oNodEntry.Attributes.Item(1).nodeValue
				sParCode=oNodEntry.Attributes.Item(3).nodeValue
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sPurInvNo=oNodEntry.Attributes.Item(0).nodeValue
				sPurInvDt=oNodEntry.Attributes.Item(1).nodeValue
				sAppTy = oNodEntry.Attributes.Item(3).nodeValue
				sAppBy = oNodEntry.Attributes.Item(4).nodeValue
			end if
			if oNodEntry.nodeName="SalesType" then
				sSalType=oNodEntry.Attributes.Item(0).nodeValue
			end if
		next
	end if
	if oNodTemp.nodeName="Details" then
		set oNodDeatils=oNodTemp
	end if
	if oNodTemp.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodTemp
		dTotal=oNodTemp.Attributes.Item(0).nodeValue
		'dTotal=oNodTemp.Attributes.Item(2).nodeValue
	end if
next

sQuery = "Select ToBeApprovedBy From Acc_T_VouchersForApproval Where CreatedTransNo = "&iCrTransNo&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sAppTy = "Y"
	sAppBy = objRs(0)
Else
	sAppTy = "N"
	sAppBy = "I"
End IF
objRs.Close

IF CStr(sAppTy) = "Y" Then
	sVouStatus="010101" 'Waiting For Approval
Else
	sVouStatus="010105" 'Waiting For Accounting 
End IF

sTransType="CNR"

sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

dim sAccHeadCode,iTemp
con.BeginTrans

iTransNo = iCrTransNo
'********************* Inserting The Data to the History Tables ***************************
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


sQuery = "INSERT INTO Acc_T_HistoryCreatedVoucherTaxDet (AmendmentNo, CreatedTransNo, AccountHead, "&_
		 "TaxCode, TaxCategoryCode, TaxEntryNo, InvoiceType, FormNumber, TaxPercentage, TaxAmount, "&_
		 "TransCrDrIndication, TransactTaxAmount) "&_
		 "Select "&iAmdendmentNo&", CreatedTransNo, AccountHead, TaxCode, TaxCategoryCode, TaxEntryNo, InvoiceType, "&_
		 "FormNumber, TaxPercentage, TaxAmount, TransCrDrIndication, TransactTaxAmount "&_
		 "FROM Acc_T_CreatedVoucherTaxDet Where CreatedTransNo = "&iTransNo

Con.Execute (sQuery)
'********************* End of Insertion The Data to the History Tables *********************

'********************* Deletion of Old Data ***********************************************

sQuery="delete Acc_T_VouchersForApproval where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CretedVoucherAHDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherCCDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherTaxDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery = "delete Acc_T_CreatedPayables where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
con.execute(sQuery)

'Response.Write sQuery &"<br><br>"
'********************* End of Deletion of Old Data ****************************************




sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
sQuery=sQuery&",'Sales Return ','SR','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

Response.Write "<p>"&sQuery &"<br><br>"
con.execute(sQuery)

for each EntryNode in oNodDeatils.childNodes
dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc

	sEntryno=EntryNode.Attributes.Item(0).nodeValue
	sAmount=EntryNode.Attributes.Item(2).nodeValue
	sItemDesc=replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
	dQty=EntryNode.Attributes.Item(3).nodeValue
	sUOM=EntryNode.Attributes.Item(4).nodeValue
	dBasicAmount="0"
	dRate=EntryNode.Attributes.Item(6).nodeValue
	dDisPer="0"
	dDisAmount="0"
	iItemCode = EntryNode.getAttribute("ItemCode")
	iClassCode = EntryNode.getAttribute("ClassCode")
	dDisPer = EntryNode.getAttribute("DisPer")
	dDisAmount = EntryNode.getAttribute("DisAmount")

	sEntryType="D"
	sNarration=""
'if	CDbl(sAmount) > 0 then
	for each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
				sAccCode=HeaderNode.Attributes.Item(0).nodeValue
				sAccType=HeaderNode.Attributes.Item(4).nodeValue
		end if 'End of Check for Account head Node
	next 'End of Entry Node Loop

	sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
	sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
	sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
	sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,ItemCode,ClassificationCode) values ("
	sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
	sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
	sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
	sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","& iItemCode &","& iClassCode &")"
    Response.Write "<p>"&sQuery &"<br><br>"
	con.execute(sQuery)
'End IF	
	
next'End of Voucher Node Loop

'-----------------------------TAX TABLE UPDATE------------------
iSno=1
dim sTaxMode
for each EntryNode in oNodTaxRoot.childNodes
	iCatCode=EntryNode.Attributes.Item(0).nodeValue
	iTaxCode=EntryNode.Attributes.Item(1).nodeValue
	sTaxMode=EntryNode.Attributes.Item(2).nodeValue
	dTaxPer=EntryNode.Attributes.Item(4).nodeValue
	dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
	sAccCode=EntryNode.Attributes.Item(6).nodeValue
	
	if	CDbl(dTaxAmount) > 0 then
		sEntryType = "D"
	else
		sEntryType = "C"
		'dTaxAmount=dTaxAmount*-1
	End if
	dTaxAmount = Abs(dTaxAmount)
	'if	CDbl(dTaxAmount) > 0 then
		if sTaxMode="P" then
			sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
				"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
				""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
			con.execute(sQuery)
		else
			sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
				"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
				""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
			con.execute(sQuery)
		end if
	'end if
	Response.Write "<p>"&sQuery &"<br><br>"


	iSno=cint(iSno)+1
Next

'----TO CODE FOR TAX FORM DEATILS----------------
Dim sPayIns,iRecNo
'----------------------------- RECEIVABLE TABLE UPDATE------------------
	sQuery = "select isnull(ReceivableNumber,0) from Acc_T_CreatedReceivables where CreatedTransNo = "&ActualTransNo&" "&_
			 "And AmountReceivable > AmountReceived and AmountReceivable - AmountReceived >= "&dTotal&" "
			 
	Response.Write "<p>"&sQuery &"<br><br>"
	objRs.open sQuery,con
	
	If Not objRs.EOF Then
		iPayableNo = objRs(0)
	Else
		iPayableNo = 0
	End If
	objRs.Close
'
	if CDbl(iPayableNo) <> 0 THEN
		sPayIns = "N"
		iRecNo = iPayableNo
		sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, ReceivedOn, "&_
				 "AmountReceived) VALUES ("&iPayableNo&", "&iCrTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dTotal&") "
				 
		Response.Write "<p>"&sQuery &"<br><br>"
				 
		' Date		:	18/11/2004
		' Reason	:	Modified for Party settelment change Requested (To treate CR Note as Seprate Adj entry)
		con.execute(sQuery)
		
		'sQuery = "insert into Acc_T_RcvblAdjustmentDetails (ReceivableNumber, RecdByTransactionNo, ReceivedOn, AmountReceived) "&_
		'		 "values("&iPayableNo&","&iTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dTotal&")"
				
		'con.execute(sQuery)
		
	else
'----------------------------- CREATE PAYABLE IF RECEIVABLE DETAILS DOESNOT EXIST ------------------
		sPayIns = "Y"
		
		sQuery = "select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close

		sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber,CreatedTransNo, OUDefinitionID, "&_
				"VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
				"PartyBillDate, AmountPayable, AmountPaid) values("&iPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
				"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0)"
				
		Response.Write "<p>"&sQuery &"<br><br>"
		
		con.execute(sQuery)
	end if
	
'////////////////////////////////////////////////////////////////////////////////////////////

IF Cstr(sAppBy) <> "I" Then
	sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
		 	"Values("&iCrTransNo&",1,"&sAppBy&" ) "
Response.Write "<p>"&sQuery &"<br><br>"
	Con.Execute sQuery
End IF
'/////////////////////////////////////// Approval Entry ///////////////////////////////////// 

'=============================================================================================
'Added On		:	29/11/2004
'Reason			:	To adjust the Credit note with the debit note 
'=============================================================================================
Dim sExp,TempNode,iAccPayNo,iCrRecNo,dAdvAmt
'This is to check wheather the Created Payables tabel is inserted or not
IF CStr(sPayIns) = "Y"  Then
	dAdjAmtTotal = 0
	sExp = "//AdvanceDetails/Advance[@AmountToAdj!=0]"
	Set TempNode = oNodRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCounter = 0 To TempNode.length - 1
			iAccPayNo = TempNode.ITem(iCounter).Attributes.getNamedItem("CreatedTransNo").value
			dAdvAmt = Trim(TempNode.ITem(iCounter).Attributes.getNamedItem("AmountToAdj").value)
			dAdvAmt = CDbl(dAdvAmt)
			dAdjAmtTotal = Cdbl(dAdjAmtTotal + dAdvAmt)
			
			
			sQuery = "Select DRCreatedReceivable From Acc_T_Receivables Where ReceivableNumber = "&iAccPayNo&" "
				
			Response.Write sQuery &"<br><br>"
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				iCrRecNo = objRs(0)
			Else
				iCrRecNo = 0
			End IF
			objRs.Close
				
			sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
					 "ReceivedOn,AmountReceived) Values ("&iCrRecNo&","&iCrTransNo&","&_
					 "getdate(),"&dAdvAmt&")"
						 
			Response.Write "<p>"&sQuery &"<br><br>"
			Con.Execute sQuery
		Next
		
		'The adjusted total debit note will get adjusted with the Created Credit Note.
		sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, PaidOn, "&_
				 "AmountPaid) VALUES ("&iPayableNo&", "&iCrTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dAdjAmtTotal&") "
				 
		Con.Execute sQuery
		Response.Write "<p>"&sQuery &"<br><br>"
		
	End IF
End IF
'Con.RollBacktrans
'Response.End 
con.CommitTrans




if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	
	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
	
	if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher Entry_CNAmd_"&Session.SessionID&".xml")) then
		objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher Entry_CNAmd_"&Session.SessionID&".xml"))
	End IF
	if objfs.FileExists(Server.MapPath("../temp/transaction/Sal_Return_Voucher_xml_"&Session.SessionID&".xml")) then
		objfs.DeleteFile(Server.MapPath("../temp/transaction/Sal_Return_Voucher_xml_"&Session.SessionID&".xml"))
	End IF
	
	
	IF CStr(sCallTy) = "OINV" Then
		Response.Redirect ("VouCNOtherInvDisplay.asp?TransNo="&iCrTransNo)
	Else
		Response.Redirect ("VouCNSalReturnDisplay.asp?TransNo="&iCrTransNo)
	End IF
	
		
end if
%>