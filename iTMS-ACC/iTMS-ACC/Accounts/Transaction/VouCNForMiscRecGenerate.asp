<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNPurInvGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 18, 2003
	'Modified By				:	Manohar Prabhu R
	'Modified On				:	April 14, 2004
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
<!--#include virtual="/include/NoSeries.asp"-->
<%

'XML DOM Variables
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot,oNodTemp,sQuery
dim sParCode,sParSubType,sParType
dim EntryNode,HeaderNode,nodANL,newElem
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,sVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iCrTransNo
dim sDocType,sVouStatus,sAppBy,sAppTy
dim iSeriesNo,iSeriesCode,sPurchaseType, ActualTransNo
Dim sExp,NarrNode,sCallFrom
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
Dim TaxCalNode,dTaxValNoAcc,iCtr,dEachItmVal,dEachItm,dNoAccTotVal
dim dAdjAmtTotal,dDiffVal,sDrAgain,sPayAt
dim sAccHeadCode,iTemp,iMiscTransNo,sInvoiceNo,sRetVal,iMiscInvNo,iCRDRNoteEntryNo
Dim rsTemp

set objRs  = server.CreateObject("adodb.recordset")
set rsTemp  = server.CreateObject("adodb.recordset")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
sInvoiceNo = Request.Form("hdTransNo")
sRetVal = GetVouchXML(sInvoiceNo)
oDOM.Load server.MapPath(sRetVal)

sVouStatus="010105" 'Crearted For Accounting to be Approved

sVouType = "C"
sDrAgain = Request.Form("selCrAgain")

sCallFrom = Request.Form("hCallFromDebit")

sAppTy = Request.Form("optApprove")
sAppBy = Request.Form("selUserId")
sNarration = Request.Form("txtNarration")
sVoucDate = Request.Form("hVouDate")

if Trim(sNarration)="" or IsNull(sNarration) then sNarration = "NULL"

if sCallFrom = "GJ" then
    sVouCode = "08"
    sTransType="GJR"
else
    sVouCode="07"
    sTransType="CNR"
end if


set oNodRoot = oDOM.documentElement
if oNodRoot.hasChildNodes() then
    sOrgId = oNodRoot.getAttribute("UnitNo")
    For each oNodEntry in oNodRoot.childNodes
        if oNodEntry.nodeName="Entry" then
            dTotal = cdbl(dTotal) + cdbl(oNodEntry.getAttribute("Amount"))
        end if
    Next
end if


IF CStr(sAppTy) = "Y" Then
	sVouStatus = "010101"
Else
	sVouStatus = "010105"
End IF

dTaxValNoAcc = 0
sBookNo = 1


    sQuery="Select CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
	       "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
    Response.Write "<p>"& sQuery
    objRs.open sQuery,con
    if not objRs.EOF then
	    iSeriesNo=objRs(0)
	    iSeriesCode=objRs(1)
    end if
    objRs.close()
    
    sVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
    
    con.BeginTrans

sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
objRs.open sQuery,con
if not objRs.EOF then 
	iCrTransNo=objRs(0)
end if
objRs.Close


sQuery = "Select PartyType,PartySubType,PartyCode,AppRefNo,MiscTransNo from Acc_T_MiscPymtRequestHeader where ApplicationCode = 3 and ReceiptNo = "& sInvoiceNo 
objRs.Open sQuery,con
if not objRs.EOF then
    sParType = objRs(0)
    sParSubType = objRs(1)
    sParCode = objRs(2)
    iMiscInvNo = objRs(3)    
    iMiscTransNo = objRs(4)
end if
objRs.Close 
if Trim(iMiscInvNo)<>"" then
    sQuery = "Select InvoiceNumber,Convert(varchar,InvoiceDate,103) from Sal_T_InvoiceHeader where SaleTransactionNo = "& iMiscInvNo 
    objRs.Open sQuery,con
    if not objRs.EOF then
        sPurInvNo = objRs(0)
        sPurInvDt = objRs(1)
    end if
    objrs.Close 
end if


sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus, BankInstrumentNo,PayableAt,PurchaseBillType) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
sQuery=sQuery&",'Sales Invoice ','OS','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"', "&sInvoiceNo&",'"&sPayAt&"','"&sDrAgain&"')"

Response.Write sQuery &"<br><br>"

con.execute(sQuery)

iCtr = 1

if Trim(sNarration)<>"NULL" then sNarration = Pack(sNarration)

for each EntryNode in oNodRoot.childNodes
	dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc

	sEntryno=EntryNode.getAttribute("No")
	sAmount=EntryNode.getAttribute("Amount")
	dQty = 0
	dRate = 0
	dDisPer = 0
	dDisAmount = 0 
	For each oNodTemp in EntryNode.childNodes
	    if oNodTemp.nodeName="Narration" then
	        sItemDesc = oNodTemp.text
	    end if
	Next
	
	dBasicAmount="0"
	sEntryType = "C"
	
	Response.Write sAmount &"<br><br><br>"

	if	CDbl(sAmount) > 0 then
		for each HeaderNode in EntryNode.childNodes
			if HeaderNode.nodeName="AccHead" then
					sAccCode=HeaderNode.Attributes.Item(0).nodeValue
					sAccType=HeaderNode.Attributes.Item(4).nodeValue
			end if 'End of Check for Account head Node
		next 'End of Entry Node Loop

		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
		sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount ) values ("
		sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
		sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery=sQuery&" "&sNarration&","&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&")"

		Response.Write sQuery &"<br><br>"
		con.execute(sQuery)

	
	end if
next'End of Voucher Node Loop

	sQuery = "select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close
	'	IF trim(sAccType) = "G" then 
'
'			sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber,CreatedTransNo, OUDefinitionID, "&_
'					"VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
'					"PartyBillDate, AmountPayable, AmountPaid) values("&iPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
'					"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
''					"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&","&dTotal&")"
'		Else
			sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber,CreatedTransNo, OUDefinitionID, "&_
					"VoucherDate, PartyType, PartySubType, PartyCode,PartyBillNumber, "&_
					"PartyBillDate, AmountPayable, AmountPaid) values("&iPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
					"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
					"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0)"
'		End IF
		Response.Write "<p>"& sQuery
		con.execute(sQuery)
	
	
	
IF Cstr(sAppBy) <> "I" Then
	sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
		 	"Values("&iCrTransNo&",1,"&sAppBy&" ) "
		 	
	Response.Write sQuery &"<br><br>"

	Con.Execute sQuery
End IF

sQuery = "Select isNull(Max(CRDRNoteEntryNo),0)+1 from ACC_T_CreatedVoucherCRDRNotes"
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    iCRDRNoteEntryNo = rsTemp(0)    
end if
rsTemp.Close 

sQuery = "Insert into ACC_T_CreatedVoucherCRDRNotes (CRDRNoteEntryNo,CROrDRNote,CreatedTransNo,RefType,RefNumber)"&_
         " values("& iCRDRNoteEntryNo &",'CNR',"& iCrTransNo &",'MI',"& iMiscTransNo &")"
Response.Write "<p>"& sQuery
con.execute sQuery


if Trim(sInvoiceNo)<>"" then
   sQuery = " Update Acc_T_MiscPymtRequestHeader set AdjustmentStatus='Y' where ReceiptNo = "& sInvoiceNo
   Response.Write "<p>"& sQuery
   con.execute sQuery
end if 'sInvoiceNo

'----To code for Tax Form deatils----------------

if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	
	Con.RollbackTrans
	Response.End
	con.CommitTrans
	
	Response.Redirect ("CreditNoteToCreate.asp")
end if

%>
