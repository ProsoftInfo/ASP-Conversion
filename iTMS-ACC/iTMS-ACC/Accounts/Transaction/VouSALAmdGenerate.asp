<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouSALAmdGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February  21, 2003
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
dim EntryNode,HeaderNode,nodANL,newElem,sExp
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus
dim iSeriesNo,iSeriesCode
dim sSalType,sRefernceNo,iCounter
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sSalInvNo,sSalInvDt,dRoundOff,sRoundoffHead
dim iCrTransNo,sCrVouNo,iCrSeriesNo,iCrSeriesCode
Dim sAppTy,sAppBy,sStr,TempNode,iAmdendmentNo,objfs,dRndOff
Dim dNoofPack,dPackTy,dRatePer


dim dAdjAmtTotal
sVouStatus="010101" 'Crearted For Approval
sVouCode="05"
sVouType="D"
sTransType="SJR"

Set objfs = CreateObject("Scripting.FileSystemObject")
set objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
if objfs.FileExists(server.MapPath("../../Sales/xmldata/General.xml")) then
    oDOM.Load server.MapPath("../../Sales/xmldata/General.xml")
    Set oNodRoot = oDOM.documentElement
    sExp = "//ROUNDOFF"
    Set oNodEntry = oNodRoot.Selectnodes(sExp)
    If oNodEntry.Length > 0 then
	    sRoundoffHead = oNodEntry.Item(0).Attributes.Item(0).nodevalue
    else
	    sRoundoffHead = "0"
    end if
else
    sRoundoffHead = "0"
end if

oDOM.load server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")

set oNodRoot=oDOM.documentElement

sStr = "//Voucher"
Set TempNode = oNodRoot.selectNodes(sStr)
IF TempNode.length <> 0 Then
	iCrTransNo = TempNode.Item(0).Attributes.Item(0).value
End IF

for each oNodTemp in oNodRoot.childNodes
	if oNodTemp.nodeName="Header" then
		for Each oNodEntry in  oNodTemp.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Book" then
				sBookNo=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="SalesType" then
				sSalType=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Party" then
				sParType=oNodEntry.Attributes.Item(0).nodeValue
				sParSubType=oNodEntry.Attributes.Item(1).nodeValue
				sParCode=oNodEntry.Attributes.Item(3).nodeValue
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sSalInvNo=oNodEntry.Attributes.Item(0).nodeValue
				sSalInvDt=oNodEntry.Attributes.Item(1).nodeValue
				sRefernceNo=oNodEntry.Attributes.Item(2).nodeValue
				sAppTy = oNodEntry.Attributes.Item(2).nodeValue
				sAppBy = oNodEntry.Attributes.Item(4).nodeValue

			end if

		next
	end if
	if oNodTemp.nodeName="AgentDetails" then
		set oNodAgent=oNodTemp
	end if

	if oNodTemp.nodeName="Details" then
		set oNodDeatils=oNodTemp
	end if
	if oNodTemp.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodTemp
		dTotal=oNodTemp.Attributes.Item(2).nodeValue
		dRoundOff=oNodTemp.Attributes.Item(3).nodeValue
	end if
	if oNodTemp.nodeName="AdvanceDetails" then
		set oNodAdvRoot=oNodTemp
	end if
next

sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

Set newElem = oDOM.CreateElement("Tax")
	newElem.SetAttribute "CatCode","0"
	newElem.SetAttribute "TaxCode","0"
	newElem.SetAttribute "TaxMode","0"
	newElem.SetAttribute "TaxFormula","0"
	newElem.SetAttribute "TaxValue","0"
	newElem.SetAttribute "TaxAmount",Cdbl(dRoundOff)
	newElem.SetAttribute "AccHead",sRoundoffHead
	newElem.Text = "ROUND OFF"
oNodTaxRoot.Appendchild newElem


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

sQuery = "INSERT INTO Acc_T_HisCreateReceivables (AmendmentNo, ReceivableNumber, CreatedTransNo, "&_
		 "OUDefinitionID, VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber, "&_
		 "PartyInvoiceDate, AmountReceivable, AmountReceived, Narration) "&_
		 "Select "&iAmdendmentNo&", ReceivableNumber, CreatedTransNo, OUDefinitionID, VoucherDate, PartyType, "&_
		 "PartySubType, PartyCode, PartyInvoiceNumber, PartyInvoiceDate,AmountReceivable, "&_
		 "AmountReceived, Narration FROM Acc_T_CreatedReceivables Where CreatedTransNo = "&iTransNo&" "

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

sQuery = "Delete Acc_T_CreatedVoucherTaxDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery = "Delete Acc_T_CreatedReceivables where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery = "Delete Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery = "Delete Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery = "Select CreatedAdvanceNo,AmountAdjusted From ACC_T_CreatedAdvanceAdj  "&_
		 "Where CreatedTransNo = "&iTransNo&" "
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With
Set objRs.ActiveConnection = Nothing
Do While Not objRs.EOF
	sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = AdvanceAdjusted - "&objRs(1)&" "&_
			 "Where CreatedAdvanceNo = "&objRs(0)&" "
	
	Response.Write sQuery &"<br><br>"		 
	Con.Execute sQuery
	objRs.MoveNext
Loop
objRs.Close

sQuery = "Delete From ACC_T_CreatedAdvanceAdj Where CreatedTransNo = "&iTransNo
con.execute(sQuery)

Response.Write sQuery &"<br><br>"

sQuery = "Delete Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
con.execute(sQuery)


'********************* End of Deletion of Old Data ****************************************

sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
objRs.open sQuery,con
	iCrTransNo=objRs(0)
objRs.Close

sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sSalInvNo&"',convert(datetime,'"&sSalInvDt&"',103),"&dTotal
sQuery=sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

'Response.Write sQuery& "<BR>"
con.execute(sQuery)

for each EntryNode in oNodDeatils.childNodes
dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc,dItemCode,dClassCode

	sEntryno=EntryNode.getAttribute("No")
	sAmount=EntryNode.getAttribute("Amount") 
	sItemDesc=Replace(EntryNode.getAttribute("PayTo"),"'","''")
	dQty=EntryNode.getAttribute("Qty")
	sUOM=EntryNode.getAttribute("UOM")
	dBasicAmount=EntryNode.getAttribute("ActValue")
	dRate=EntryNode.getAttribute("Rate")
	dDisPer=EntryNode.getAttribute("DisPer")
	dDisAmount=EntryNode.getAttribute("DisAmount")
	dRndOff =EntryNode.getAttribute("RndOff")
	dNoofPack = EntryNode.getAttribute("NoofPack")
	dPackTy = EntryNode.getAttribute("PackType")
	dRatePer = EntryNode.getAttribute("RatePer")
	dItemCode = EntryNode.getAttribute("ItemCode")
	dClassCode = EntryNode.getAttribute("ClassCode")

	sEntryType="C"
	sNarration=""

	for each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
				sAccCode=HeaderNode.Attributes.Item(0).nodeValue
				sAccType=HeaderNode.Attributes.Item(4).nodeValue
		end if 'End of Check for Account head Node
	next 'End of Entry Node Loop
	
	IF CStr(dNoofPack) = "" Then
		dNoofPack = 0
	End IF

		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
		sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,RoundOffvalue,NoofPack,PackCode,RatePer,ItemCode,ClassificationCode ) values ("
		sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
		sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","&dRndOff&","&dNoofPack&","&dPackTy&","&dRatePer&","&dItemCode&", "&dClassCode&" )"
Response.Write sQuery& "<BR>"
con.execute(sQuery)

dim sCCGroup,sAddCode,sAddRatio,sAddAmount

for each HeaderNode in EntryNode.childNodes
	if 	HeaderNode.nodeName="CostCenter" then
		for each  nodANL in HeaderNode.childNodes
			sAddCode=nodANL.Attributes.Item(0).nodeValue
			sAddRatio=nodANL.Attributes.Item(3).nodeValue
			sAddAmount=nodANL.Attributes.Item(4).nodeValue

			sQuery="INSERT INTO Acc_T_CreatedVoucherCCDet(CreatedTransNo, VoucherEntryNumber, AccountingUnit,"&_
				" AccUnitAccountHead,AccUnitCCHead,"&_
				"CCRatioPercent, CCRatioAmount)"&_
				" VALUES("& iCrTransNo& ","& sEntryno& ",'"&sOrgId&"',"&sAccCode&","&_
				" "& sAddCode &","& sAddRatio &"," & sAddAmount & ")"
			'Response.Write sQuery& "<BR>"
			con.execute(sQuery)

		next
	end if 'End of Check for Cost Center Node
	if 	HeaderNode.nodeName="Analytical" then
		for each  nodANL in HeaderNode.childNodes
			sAddCode=nodANL.Attributes.Item(0).nodeValue
			sAddRatio=nodANL.Attributes.Item(3).nodeValue
			sAddAmount=nodANL.Attributes.Item(4).nodeValue

			sQuery="INSERT INTO Acc_T_CretedVoucherAHDet(CreatedTransNo, VoucherEntryNumber, AccountingUnit, "&_
				"AccUnitAccountHead, AccUnitAnalyticalCode,"&_
				"RatioPercentage, RatioAmount)"&_
				" VALUES("& iCrTransNo& ","& sEntryno& ",'"&sOrgId&"',"&sAccCode&","&_
				""&sAddCode&","&sAddRatio&","&sAddAmount&")"
			'Response.Write sQuery& "<BR>"
			con.execute(sQuery)

		next
	end if 'End of Check for Analytical Node
next 'End of Entry Node Loop

next'End of Voucher Node Loop

iSno=1
dim sTaxMode
for each EntryNode in oNodTaxRoot.childNodes
	iCatCode=EntryNode.Attributes.Item(0).nodeValue
	iTaxCode=EntryNode.Attributes.Item(1).nodeValue
	sTaxMode=EntryNode.Attributes.Item(2).nodeValue
	dTaxPer=EntryNode.Attributes.Item(4).nodeValue
	dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
	sAccCode=EntryNode.Attributes.Item(6).nodeValue
	if dTaxAmount >=0 then
		sEntryType = "C"
	else
		sEntryType = "D"
		dTaxAmount=dTaxAmount*-1
	End if
	if sTaxMode="P" then
		sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
			"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
			""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
	else
		sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
			"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
			""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
	end if
'Response.Write sQuery& "<BR>"
con.execute(sQuery)
iSno=cint(iSno)+1
Next


dim iAdvTranNo,dAdvAmount,iCrAdvNo,sAdvNarr
dAdjAmtTotal=0
if  IsObject(oNodAdvRoot) then
	for each EntryNode in oNodAdvRoot.childNodes
		iAdvTranNo=EntryNode.Attributes.Item(0).nodeValue
		dAdvAmount=EntryNode.Attributes.Item(5).nodeValue
		iCrAdvNo = EntryNode.Attributes.getNamedItem("AdvNo").nodeValue
		dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
		
		IF CDbl(dAdvAmount) >0 Then

			sQuery = "Update Acc_T_CreatedAdvances set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
					 " Where CreatedAdvanceNo = "&iCrAdvno
					 
			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)
		
			sAdvNarr = "Sales Invoice No: "&sSalInvNo&" DT: " & sVoucDate
			sQuery = "INSERT INTO ACC_T_CreatedAdvanceAdj (CreatedAdvanceNo, CreatedTransNo,  "&_
					 "AdjustedOn, AmountAdjusted, Narration)  "&_
					 "VALUES ("&iCrAdvno&", "&iCrTransNo&", Convert(Datetime,'"&sSalInvDt&"',103), "&dAdvAmount&", '"&sAdvNarr&"') "
						 
			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)
		End IF
			
	Next
end if
dim iCrPayable,sPayNarration
if CDbl(dTotal)>CDbl(dAdjAmtTotal) then
	sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_CreatedReceivables"
	objRs.open sQuery,con
		iCrPayable=objRs(0)
	objRs.Close

sPayNarration="SALE INV NO:"&sSalInvNo&" Dt:"&sSalInvDt

	sQuery="INSERT INTO Acc_T_CreatedReceivables(ReceivableNumber, CreatedTransNo, OUDefinitionID,"&_
			"VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber,"&_
			" PartyInvoiceDate, AmountReceivable, AmountReceived,Narration)values("&iCrPayable&","&iCrTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sSalInvNo&"',"&_
			"convert(datetime,'"&sSalInvDt&"',103),"&dTotal&","&dAdjAmtTotal&",'"&sPayNarration&"')"
			
	Response.Write sQuery &"<br><br>"
	
	con.execute(sQuery)


IF CStr(sAppBy) = "Y" Then
	sAppBy = GetUserID()
End IF

sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
		 "Values("&iCrTransNo&",1,"&sAppBy&" ) "
Con.Execute sQuery

end if
'Response.Write "<BR><BR>"

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
	

	Set newElem  = oDOM.createAttribute("CreatedTransNo")
	newElem.value = iCrTransNo
	oNodRoot.setAttributeNode(newElem)

	Set newElem  = oDOM.createAttribute("CreatedVouNo")
	newElem.value = sSalInvNo
	oNodRoot.setAttributeNode(newElem)

	Set newElem  = oDOM.createAttribute("TransNo")
	newElem.value = iTransNo
	oNodRoot.setAttributeNode(newElem)

	Set newElem  = oDOM.createAttribute("VouNo")
	newElem.value = sSalInvNo
	oNodRoot.setAttributeNode(newElem)

	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
	if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")) then
		objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"))
	End IF

	Response.Redirect ("VouSALDisplay.asp?TransNo="&iCrTransNo&"&CallFrm=A")

end if
%>
