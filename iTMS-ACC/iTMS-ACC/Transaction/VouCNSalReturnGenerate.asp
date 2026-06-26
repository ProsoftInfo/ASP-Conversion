<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNSalReturnGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/NoSeries.asp"-->
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
dim dAdjAmtTotal,sAppTy,sAppBy,iCrTransNo,objfs,sCallTy,dTotalVouchVal
Dim iSalInvoiceNo,iSalInvoiceReturnNo,iCRDRNoteEntryNo
Dim rsTemp,iItemCode,iClassCode

sAppTy = Request.Form("optApprove")
sCallTy = Request.Form("hCallType")

IF CStr(sAppTy) = "Y" Then
	sVouStatus="010101" 'Waiting For Approval
Else
	sVouStatus="010105" 'Waiting For Accounting
End IF

sVouCode="07"
sVouType="C"

ActualTransNo = Request("hdTransNo")
sCallTy = Request.Form("hCallType")

set objRs  = server.CreateObject("adodb.recordset")
set rsTemp  = server.CreateObject("adodb.recordset")
Set objfs = CreateObject("Scripting.FileSystemObject")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.load server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")
set oNodRoot=oDOM.documentElement

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
				sAppTy = oNodEntry.getAttribute("Approval")
				sAppBy = oNodEntry.getAttribute("Approver")
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
		'dTotal=oNodTemp.Attributes.Item(0).nodeValue
		dTotalVouchVal=Trim(oNodTemp.Attributes.Item(0).nodeValue)
		dTotal=oNodTemp.Attributes.Item(0).nodeValue
	end if
	if oNodTemp.nodeName="SalesInvoiceEntry" then
        iSalInvoiceNo = oNodTemp.getAttribute("InvoiceNo")
        Response.Write "<p> Invoice No = "& iSalInvoiceNo
    end if
next

sTransType="CNR"

sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

dim sAccHeadCode,iTemp

sQuery="select CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
	"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo

objRs.open sQuery,con
if not objRs.EOF then
	iSeriesNo=objRs(0)
	iSeriesCode=objRs(1)
end if
objRs.close()

con.BeginTrans

iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)

sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
objRs.open sQuery,con
	iCrTransNo=objRs(0)
objRs.Close

sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotalVouchVal
sQuery=sQuery&",'Sales Return ','SR','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

Response.Write "<P>"& sQuery &"<br><br>"
con.execute(sQuery)



sQuery = " Select SalesReturnNo from Sal_T_SalesReturnHeader where  SaleTransactionNo = "& iSalInvoiceNo  &" order by SalesReturnNo desc"
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    iSalInvoiceReturnNo = rsTemp(0)
end if
rsTemp.Close 

sQuery = "Select isNull(Max(CRDRNoteEntryNo),0)+1 from ACC_T_CreatedVoucherCRDRNotes"
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    iCRDRNoteEntryNo = rsTemp(0)    
end if
rsTemp.Close 

sQuery = "Insert into ACC_T_CreatedVoucherCRDRNotes (CRDRNoteEntryNo,CROrDRNote,CreatedTransNo,RefType,RefNumber)"&_
         " values("& iCRDRNoteEntryNo &",'CNR',"& iCrTransNo &",'SR',"& iSalInvoiceReturnNo  &")"
Response.Write "<p>"& sQuery
con.execute sQuery

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
Response.Write "<P>"& sQuery &"<br><br>"
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
	End if

	dTaxAmount = Abs(dTaxAmount)

'	if	CDbl(dTaxAmount) > 0 then
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
'	end if

Response.Write "<P>"& sQuery &"<br><br>"

	iSno=cint(iSno)+1
Next

'----TO CODE FOR TAX FORM DEATILS----------------
Dim sPayIns,iRecNo
'----------------------------- RECEIVABLE TABLE UPDATE------------------
	sQuery = "select isnull(ReceivableNumber,0) from Acc_T_CreatedReceivables where CreatedTransNo = "&ActualTransNo&" "&_
			 "And AmountReceivable > AmountReceived and AmountReceivable - AmountReceived >= "&dTotal&" "

	Response.Write "<P>"& sQuery &"<br><br>"
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

		Response.Write "<P>"& sQuery &"<br><br>"

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

		Response.Write "<P>"& sQuery &"<br><br>"

		con.execute(sQuery)
	end if

'////////////////////////////////////////////////////////////////////////////////////////////

IF Cstr(sAppBy) <> "I" Then
	sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
		 	"Values("&iCrTransNo&",1,"&sAppBy&" ) "
Response.Write "<P>"& sQuery &"<br><br>"
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
	Response.Write sExp
	Set TempNode = oNodRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCounter = 0 To TempNode.length - 1
			iAccPayNo = TempNode.ITem(iCounter).Attributes.getNamedItem("CreatedTransNo").value
			dAdvAmt = Trim(TempNode.ITem(iCounter).Attributes.getNamedItem("AmountToAdj").value)
			dAdvAmt = CDbl(dAdvAmt)
			dAdjAmtTotal = Cdbl(dAdjAmtTotal + dAdvAmt)


			sQuery = "Select DRCreatedReceivable From Acc_T_Receivables Where ReceivableNumber = "&iAccPayNo&" "

			Response.Write "<P>"& sQuery &"<br><br>"
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

			Response.Write "<P>"& sQuery &"<br><br>"
			Con.Execute sQuery
		Next

		'The adjusted total debit note will get adjusted with the Created Credit Note.
		sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, PaidOn, "&_
				 "AmountPaid) VALUES ("&iPayableNo&", "&iCrTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dAdjAmtTotal&") "

		Con.Execute sQuery
		Response.Write "<P>"& sQuery &"<br><br>"

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
	'con.CommitTrans

	oNodRoot.Attributes.Item(0).nodeValue= iCrTransNo
	oNodRoot.Attributes.Item(1).nodeValue= iVouNo

	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")

	if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")) then
		objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml"))
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
