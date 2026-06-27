<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNPurInvAmdUpdate.asp
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
dim sDocType,sVouStatus,sAppBy,sAppTy,sExp,NarrNode,sDrAgain
dim iSeriesNo,iSeriesCode,sPurchaseType, ActualTransNo,sDispType,sPayAt

dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
Dim TaxCalNode,iCtr,dTaxValNoAcc,dEachItm,dEachItmVal,dDiffVal

dim dAdjAmtTotal,iItemCode,iClassCode,objfs
sVouStatus="010105" 'Crearted For Accounting to be Approved
sVouCode="06"
'sVouType = "D"
sVouType = "D"

ActualTransNo = Request("hdTransNo")
iCrTransNo = Request("hCrTransNo")
sDispType = Request("SelCrAgain")

set objRs  = server.CreateObject("adodb.recordset")
Set objfs = CreateObject("Scripting.FileSystemObject")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.load server.MapPath("../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml")

set oNodRoot=oDOM.documentElement
sVouNo = oNodRoot.Attributes.Item(1).nodeValue

for each oNodTemp in oNodRoot.childNodes
	if oNodTemp.nodeName="Header" then
		sPayAt = oNodTemp.getAttribute("PayableAt") 
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
			if oNodEntry.nodeName="PurInvoice" then
				sPurInvNo=oNodEntry.Attributes.Item(0).nodeValue
				sPurInvDt=oNodEntry.Attributes.Item(1).nodeValue
				sAppTy = oNodEntry.Attributes.getNamedItem("Approval").value
				sAppBy = oNodEntry.Attributes.getNamedItem("Approver").value
			end if
			if oNodEntry.nodeName="PurchaseType" then
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
next

sExp = "//Narration"
Set NarrNode = oNodRoot.selectNodes(sExp)
IF NarrNode.length <> 0 Then
	sNarration = NarrNode.Item(0).Text
Else
	sNarration = ""
End IF


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


sTransType="DNR"

sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue


dim sAccHeadCode,iTemp

IF CStr(sAppTy) = "Y" Then
	sVouStatus = "010101"
Else
	sVouStatus = "010105"
End IF

con.BeginTrans

'=========================== Coping the Old Value to History =================================

'=========================== Coping the Old Value to History =================================
'================================= Deleting the Old Values ===================================
sQuery = "Delete From Acc_T_VouchersForApproval Where CreatedTransNo = "&iCrTransNo&" "
Con.Execute(sQuery)
sQuery = "Delete From Acc_T_CreatedVoucherTaxDet Where CreatedTransNo = "&iCrTransNo&" "
Con.Execute(sQuery)
sQuery = "Delete From Acc_T_CreatedPybleAdjDet Where CreatedTransNo = "&iCrTransNo&" "
Con.Execute(sQuery)
sQuery = "Delete From Acc_T_CreatedReceivables Where CreatedTransNo = "&iCrTransNo&" "
Con.Execute(sQuery)
sQuery = "Delete From Acc_T_CreatedVoucherDetails Where CreatedTransNo = "&iCrTransNo&" "
Con.Execute(sQuery)
sQuery = "Delete From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iCrTransNo&" "
Con.Execute(sQuery)


'================================= Deleting the Old Values ===================================
sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus, BankInstrumentNo,PayableAt) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
sQuery=sQuery&",'Purchase Invoice ','OP','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"', "&ActualTransNo&",'"&sPayAt&"')"

Response.Write "<p>"& sQuery &"<br><br>"

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
	dDisPer=EntryNode.getAttribute("DisPer")
	dDisAmount=EntryNode.getAttribute("DisAmount")

	sEntryType = "C"
	'sNarration = ""
	
	

'	if	CDbl(sAmount) > 0 then
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

		Response.Write "<p>"& sQuery &"<br><br>"
		con.execute(sQuery)
'	end if
next'End of Voucher Node Loop


'--------------------------------------------PAYABLE TABLE UPDATE------------------

	sQuery = "select isnull(PayablesNumber,0) from Acc_T_CreatedPayables where CreatedTransNo = "&ActualTransNo&" "&_
			 "and AmountPayable > AmountPaid and AmountPayable - AmountPaid >= "&dTotal&" "
	
	objRs.open sQuery,con
	
	Response.Write "<p>"& sQuery &"<br><br>"
	

	If Not objRs.EOF Then
		iPayableNo = objRs(0)
	Else
		iPayableNo = 0
	End If
	objRs.Close
	
	Response.Write "<p>"& iPayableNo &"<br><br>"

	if cdbl(iPayableNo) <> 0 Then
		sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, PaidOn, AmountPaid) "&_
				 "VALUES ("&iPayableNo&", "&iCrTransNo&", convert(datetime,'"&sVoucDate&"',103), "&dTotal&") "
				 
		Response.Write "<p>"& sQuery &"<br><br>"
		
		con.execute(sQuery)
'----------------------------- CREATE RECEIVABLE IF PAYABLE DETAILS DOESNOT EXIST ------------------
	else
		sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_CreatedReceivables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close

		sQuery="INSERT INTO Acc_T_CreatedReceivables(ReceivableNumber, CreatedTransNo, OUDefinitionID,"&_
			"VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber,"&_
			" PartyInvoiceDate, AmountReceivable, AmountReceived)values("&iPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
			"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0)"
			
			Response.Write "<p>"& sQuery &"<br><br>"
			
		con.execute(sQuery)
	end if
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
		dTaxAmount=dTaxAmount*-1
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
		Response.Write "<p>"& sQuery &"<br><br>"
	'end if
	iSno=cint(iSno)+1
Next
'------------------------------ END OF TAX NODE
IF CStr(sAppTy) = "" Then
	sAppTy = "N"
	if Trim(sAppBy)="" or IsNull(sAppBy) then sAppBy = "NULL"
End IF 

IF Cstr(sAppTy) <> "N" Then
	sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
		 	"Values("&iCrTransNo&",1,"&sAppBy&" ) "
		 	
	Response.Write "<p>"& sQuery &"<br><br>"

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
   ' Con.RollbackTrans
'	Response.End
	con.CommitTrans
	

	'oNodRoot.Attributes.Item(0).nodeValue= iCrTransNo
	'oNodRoot.Attributes.Item(1).nodeValue= sVouNo

	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
	'Response.Redirect ("VouDNPurReturnDisplay.asp?TransNo="&iCrTransNo&"&DispType="&sDispType)
	if objfs.FileExists(server.MapPath("../temp/transaction/Voucher EntryAmd_DN_"&Session.SessionID&".xml")) then
	    objfs.DeleteFile(server.MapPath("../temp/transaction/Voucher EntryAmd_DN_"&Session.SessionID&".xml"))
	end if
	if objfs.FileExists(server.MapPath("../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml")) then
    	objfs.DeleteFile(server.MapPath("../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml"))
	end if
	
	if objfs.FileExists(server.MapPath("../temp/transaction/InvDet_ForDN_"&Session.SessionID&".xml")) then
    	objfs.DeleteFile(server.MapPath("../temp/transaction/InvDet_ForDN_"&Session.SessionID&".xml"))
	end if
	Response.Redirect ("VouDNPurReturnDisplay.asp?TransNo="&iCrTransNo)
	
end if

%>
