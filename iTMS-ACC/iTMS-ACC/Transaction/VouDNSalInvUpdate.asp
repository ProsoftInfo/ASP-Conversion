<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNPurInvUpdate.asp
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/NoSeries.asp"-->
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
Dim TaxCalNode,dTaxValNoAcc,iCtr,dEachItmVal,dEachItm,dNoAccTotVal,dDiffVal
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt,sDrAgain

dim dAdjAmtTotal,iSalInvoiceNo,iCRDRNoteEntryNo,iItemCode,iClassCode
Dim rsTemp,objfs
sVouStatus="010105" 'Crearted For Accounting to be Approved
sCallFrom = Request.Form("hCallFromDebit")

if sCallFrom = "GJ" then
    sVouCode = "08"
    sTransType="GJR"
else
    sVouCode="06"
    sTransType="DNR"
end if

Response.Write "sCallFrom = "& sCallFrom
Response.Write " sVouCode = "& sVouCode
sVouType = "D"


ActualTransNo = Request("hdTransNo")
sDrAgain = Request("SelCrAgain")

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
				sAppTy = oNodEntry.Attributes.getNamedItem("Approval").value
				sAppBy = oNodEntry.Attributes.getNamedItem("Approver").value
			end if
			if oNodEntry.nodeName="SalesType" then
				sPurchaseType=oNodEntry.Attributes.Item(0).nodeValue
			end if
		next
	end if
	if oNodTemp.nodeName="Details" then
		set oNodDeatils=oNodTemp
	end if
	if oNodTemp.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodTemp
		dTotal=oNodTemp.Attributes.Item(0).nodeValue
	end if
	if oNodTemp.nodeName="SalesInvoiceEntry" then
	    iSalInvoiceNo = oNodTemp.getAttribute("InvoiceNo")
	    Response.Write "<p> Invoice No = "& iSalInvoiceNo
	end if
next

sExp = "//Narration"
Set NarrNode = oNodRoot.selectNodes(sExp)
IF NarrNode.length <> 0 Then
	sNarration = NarrNode.Item(0).text
Else
	sNarration = ""
End IF

sTransType="DNR"

'sExp = "//TaxDetails/Tax[@AccHead=0]"
'Set TaxCalNode = oNodRoot.selectNodes(sExp)
'IF TaxCalNode.length <> 0 Then
'	For iCtr = 0 To TaxCalNode.length - 1
'		dTaxValNoAcc = CDbl(dTaxValNoAcc) + CDbl(TaxCalNode.Item(iCtr).Attributes.getNamedItem("TaxAmount").Value)
'	Next
'End IF
'sExp = "//Entry"
'Set TaxCalNode = oNodRoot.selectNodes(sExp)
'dEachItm = TaxCalNode.length

'IF CDbl(dEachItm) = 1 Then
'	dEachItmVal = dTaxValNoAcc
'Else
'	dEachItmVal =  Round(CDbl(dTaxValNoAcc) / CDbl(dEachItm),0)
'	dNoAccTotVal = CDbl(dEachItmVal) * CDbl(dEachItm)
'	dDiffVal = CDbl(dTaxValNoAcc) - CDbl(dNoAccTotVal)
'End IF


sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue


dim sAccHeadCode,iTemp

IF CStr(sAppTy) = "Y" Then
	sVouStatus = "010101"
Else
	sVouStatus = "010105"
End IF


sQuery="select CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
	"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo

objRs.open sQuery,con
if not objRs.EOF then
	iSeriesNo=objRs(0)
	iSeriesCode=objRs(1)
end if
objRs.close()

con.BeginTrans

sVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)

sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
objRs.open sQuery,con
	iCrTransNo=objRs(0)
objRs.Close

sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus, BankInstrumentNo,PurchaseBillType) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
sQuery=sQuery&",'Sales ADD ','SI','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"', "&ActualTransNo&",'"&sDrAgain&"')"
Response.Write "<p>"& sQuery
con.execute(sQuery)

sQuery = "Select isNull(Max(CRDRNoteEntryNo),0)+1 from ACC_T_CreatedVoucherCRDRNotes"
rsTemp.Open sQuery,con
if not rsTemp.EOF then
    iCRDRNoteEntryNo = rsTemp(0)    
end if
rsTemp.Close 

sQuery = "Insert into ACC_T_CreatedVoucherCRDRNotes (CRDRNoteEntryNo,CROrDRNote,CreatedTransNo,RefType,RefNumber)"&_
         " values("& iCRDRNoteEntryNo &",'DNR',"& iCrTransNo &",'SI',"& iSalInvoiceNo &")"
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
	iItemCode = EntryNode.getAttribute("ItemCode")
	iClassCode= EntryNode.getAttribute("ClassCode")
	
	dDisPer=EntryNode.getAttribute("DisPer")
	dDisAmount=EntryNode.getAttribute("DisAmount")
	
	if trim(dDisPer)="" or isnull(dDisPer) then dDisPer="0"
	if trim(dDisAmount)="" or isnull(dDisAmount) then  dDisAmount="0"

	sEntryType = "C"
	'sNarration = ""
	
	sNarration = sItemDesc

	if	CDbl(sAmount) > 0 then
		for each HeaderNode in EntryNode.childNodes
			if HeaderNode.nodeName="AccHead" then
					sAccCode=HeaderNode.Attributes.Item(0).nodeValue
					sAccType=HeaderNode.Attributes.Item(4).nodeValue
			end if 'End of Check for Account head Node
		next 'End of Entry Node Loop
		
		'sAmount = CDbl(sAmount) + CDbl(dEachItmVal)
		'IF CDbl(iCtr) = 1 Then
		'	sAmount = CDbl(sAmount) + CDbl(dDiffVal)
		'	iCtr = 2
		'End IF

		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
		sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,ItemCode,ClassificationCode) values ("
		sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
		sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","& iItemCode &","& iClassCode &")"
Response.Write "<p>"& sQuery
		con.execute(sQuery)
		iItemCode ="0"
		iClassCode = "0"
	end if
next'End of Voucher Node Loop


'--------------------------------------------PAYABLE TABLE UPDATE------------------

	sQuery = "Select Max(isNull(ReceivableNumber,0))+ 1 From Acc_T_CreatedReceivables "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		iPayableNo = objRs(0)
	Else
		iPayableNo =  "0"
	End IF
	objRs.Close
	
	sQuery = "INSERT INTO Acc_T_CreatedReceivables (ReceivableNumber, CreatedTransNo, OUDefinitionID,  "&_
			 "PartyType, PartySubType, PartyCode, VoucherDate, PartyInvoiceNumber, PartyInvoiceDate,  "&_
			 "AmountReceivable, Narration) "&_
			 "VALUES ("&iPayableNo&", "&iCrTransNo&", '"&sOrgId&"', "&_
			 "'"&sParType&"', "&sParSubType&", "&sParCode&", Convert(datetime,'"&sVoucDate&"',103), "&_
			 "'"&sPurInvNo&"', Convert(datetime,'"&sPurInvDt&"',103), "&dTotal&", 'DR Note For Sales Invoice') "
			 
    Response.Write "<p>"& sQuery
	
	con.execute(sQuery)
'------------------------------------------------------------------------------------------------------------------------

'---------------------------------- FOR TAX NODE DETAILS ----------------

iSno=1
dim sTaxMode
for each EntryNode in oNodTaxRoot.childNodes
	iCatCode=EntryNode.Attributes.Item(0).nodeValue
	iTaxCode=EntryNode.Attributes.Item(1).nodeValue
	sTaxMode=EntryNode.Attributes.Item(2).nodeValue
	dTaxPer=EntryNode.Attributes.Item(4).nodeValue
	dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
	sAccCode=EntryNode.Attributes.Item(6).nodeValue

	If dTaxAmount >=0 then
		sEntryType = "C"
	else
		sEntryType = "D"
		dTaxAmount = Abs(dTaxAmount)
	End if

	'if	CDbl(dTaxAmount) > 0 and cint(sAccCode)>0  then
		if sTaxMode="P" then
			sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
				"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
				""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sPurchaseType&",NULL,"&dTaxPer&","&dTaxAmount&")"
			con.execute(sQuery)
		else
			sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
				"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
				""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sPurchaseType&",NULL,NULL,"&dTaxAmount&")"
				
			con.execute(sQuery)
		end if
Response.Write "<p>"& sQuery
		
	'end if
	iSno=cint(iSno)+1
Next
'------------------------------ END OF TAX NODE

IF Cstr(sAppTy) <> "N" Then
	sAppBy = getUserid()
	sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
		 	"Values("&iCrTransNo&",1,"&sAppBy&" ) "
		 	
	Response.Write sQuery

	Con.Execute sQuery
End IF

'----To code for Tax Form deatils----------------

if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else

'	Con.RollbackTrans
'	Response.End
	
	con.CommitTrans

	'oNodRoot.Attributes.Item(0).nodeValue= iCrTransNo
	'oNodRoot.Attributes.Item(1).nodeValue= sVouNo

	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
	'Response.Redirect ("VouDNSalInvDisplay.asp?TransNo="&iCrTransNo)
	if objfs.FileExists(server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")) then
	    objfs.DeleteFile(server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml"))
	end if
	Response.Redirect ("DEBITVOUCHERS.ASP")
end if

%>
