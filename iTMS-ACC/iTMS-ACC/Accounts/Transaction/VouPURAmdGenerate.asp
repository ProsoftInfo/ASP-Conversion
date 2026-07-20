<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouPURAmdGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu R
	'Created On					:	OCT 28, 2004
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
dim sDocType,sVouStatus
dim iSeriesNo,iSeriesCode
dim sSalType,sRefernceNo,iCounter
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
dim iCrTransNo,sCrVouNo,iCrSeriesNo,iCrSeriesCode
Dim sAppTy,sAppBy,sExp,dRoundOff,sRoundoffHead,sGroupCode,bAddFlag,Objfs
Dim CheckNode,CheckVal
Dim iAmdendmentNo,dInvWtRnd,sFormVal
dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc,dClassCode,dItemCode

dim dAdjAmtTotal
sVouStatus="010105" 'Crearted For Accounting to be Approved
sVouCode="04"
sVouType="C"
sFormVal = Request.Form("hFormVal")

set objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
IF Objfs.FileExists(server.MapPath("../../Purchase/xmldata/General.xml")) THEN
    oDOM.Load server.MapPath("../../Purchase/xmldata/General.xml")
    Set oNodRoot = oDOM.documentElement
    sExp = "//ROUNDOFF"
    Set oNodEntry = oNodRoot.Selectnodes(sExp)
    If oNodEntry.Length > 0 then
	    sRoundoffHead = oNodEntry.Item(0).Attributes.Item(0).nodevalue
    else
	    sRoundoffHead = "0"
    end if
ELSE
    sRoundoffHead = "0"
END IF

oDOM.load server.MapPath("../temp/transaction/Voucher Amd_PUR_"&Session.SessionID&".xml")

set oNodRoot=oDOM.documentElement

'Response.End

sExp = "//TaxDetails[@RoundOffValue!=0]"
Set CheckNode = oNodRoot.selectNodes(sExp)
IF CheckNode.length <> 0 Then
	CheckVal = "T"
Else
	CheckVal = "F"
End IF

iCrTransNo = oNodRoot.Attributes.Item(0).nodeValue
sCrVouNo = oNodRoot.Attributes.Item(1).nodeValue
iTransNo = iCrTransNo


for each oNodTemp in oNodRoot.childNodes
	if oNodTemp.nodeName="Header" then
		for Each oNodEntry in  oNodTemp.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Book" then
				sBookNo=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="PurchaseType" then
				sSalType=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Party" then
				sParType=oNodEntry.Attributes.Item(0).nodeValue
				sParSubType=oNodEntry.Attributes.Item(1).nodeValue
				sParCode=oNodEntry.Attributes.Item(3).nodeValue
			end if
			if oNodEntry.nodeName="PurInvoice" then
				sPurInvNo=oNodEntry.getAttribute("PurInvNo")
				sPurInvDt=oNodEntry.getAttribute("PurInvDate")
				sAppTy = oNodEntry.getAttribute("Approval")
				sAppBy = oNodEntry.getAttribute("Approver")
				sRefernceNo=sPurInvNo&"-"&sPurInvDt
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
		'dTotal=oNodTemp.Attributes.Item(0).nodeValue
		dTotal=oNodTemp.Attributes.Item(2).nodeValue
		IF CStr(CheckVal) = "T" Then
			dRoundOff=oNodTemp.Attributes.Item(3).nodeValue
		Else
			dRoundOff = 0
		End IF
	end if
	if oNodTemp.nodeName="AdvanceDetails" then
		set oNodAdvRoot=oNodTemp
	end if


next

sExp = "//Details"
Set TempNode = oNodRoot.selectnodes(sExp)
For iCounter = 0 To TempNode.length - 1
	IF TempNode.Item(iCounter).Attributes.getNamedItem("BasicValue").Value <> "" Then
		Response.Write iCounter
		set oNodDeatils = TempNode.Item(iCounter)
	End IF
Next

sTransType="PJR"

IF CStr(sAppTy) = "Y" Then
	sVouStatus = "010101" 'Voucher To be Approved"
Else
	sVouStatus = "010105" 'Voucher approved to be Accounted
End IF

sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

'Addition of Roundoff Node
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



'IF CStr(sRoundoffHead) = "0" Then
'	sExp = "//Entry"
'	Set CheckNode = oNodRoot.selectNodes(sExp)
'	IF CheckNode.length <> 0 Then
'		dInvWtRnd = CheckNode.Item(0).Attributes.Item(2).NodeValue
'		dInvWtRnd = CDbl(dInvWtRnd) + CDbl(dRoundOff)
'		CheckNode.Item(0).Attributes.Item(2).NodeValue  = dInvWtRnd
'	End IF
'End IF


dim sAccHeadCode,iTemp

con.BeginTrans

iVouNo=""


'================= Insertion For the History Tabel Starts Here ===============================
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

sQuery = "INSERT INTO Acc_T_HisCreatedPayables (AmendmentNo, PayablesNumber, CreatedTransNo,  "&_
		 "OUDefinitionID, PartyType, PartySubType, PartyCode, VoucherDate, PartyBillNumber,  "&_
		 "PartyBillDate, AmountPayable, AmountPaid, Narration) "&_
		 "SELECT "&iAmdendmentNo&", PayablesNumber, CreatedTransNo, OUDefinitionID, PartyType, PartySubType, PartyCode, "&_
		 "VoucherDate, PartyBillNumber, PartyBillDate, AmountPayable, AmountPaid, Narration "&_
		 "FROM Acc_T_CreatedPayables Where CreatedTransNo = "&iCrTransNo&" "

Response.Write sQuery

con.execute(sQuery)

sQuery = "INSERT INTO Acc_T_HistoryPybleAdjustment (PayablesNumber, PaidByTransactionNo, "&_
		 "AmendmentNo, PaidOn, AmountPaid) "&_
		 "SELECT PayablesNumber, CreatedTransNo, "&iAmdendmentNo&", PaidOn, AmountPaid FROM   "&_
		 "Acc_T_CreatedPybleAdjDet Where CreatedTransNo = "&iCrTransNo&" "

con.execute(sQuery)


sQuery = "INSERT INTO Acc_T_HistoryCreatedVoucherTaxDet (AmendmentNo, CreatedTransNo, AccountHead, "&_
		 "TaxCode, TaxCategoryCode, TaxEntryNo, InvoiceType, FormNumber, TaxPercentage, TaxAmount, "&_
		 "TransCrDrIndication, TransactTaxAmount) "&_
		 "Select "&iAmdendmentNo&", CreatedTransNo, AccountHead, TaxCode, TaxCategoryCode, TaxEntryNo, InvoiceType, "&_
		 "FormNumber, TaxPercentage, TaxAmount, TransCrDrIndication, TransactTaxAmount "&_
		 "FROM Acc_T_CreatedVoucherTaxDet Where CreatedTransNo = "&iTransNo

Con.Execute (sQuery)

'================= Insertion For the History Tabel Starts Here ===============================

'==================== Removing of Older Values From the Tabel ================================
sQuery="delete Acc_T_VouchersForApproval where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CretedVoucherAHDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherCCDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherTaxDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedPayables where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery = "Delete Acc_T_CreatedRcvbleAdjDet Where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo
con.execute(sQuery)

Response.Clear

sQuery = "Select CreatedAdvanceNo,AmountAdjusted From ACC_T_CreatedAdvanceAdj "&_
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

	'Response.Write sQuery &"<br><br>"
	Con.Execute sQuery
	objRs.MoveNext
Loop
objRs.Close

sQuery = "Delete From ACC_T_CreatedAdvanceAdj Where CreatedTransNo = "&iTransNo
'Response.Write sQuery &"<br><br>"
con.execute(sQuery)


sQuery="delete Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
'Response.Write sQuery &"<br><br>"
con.execute(sQuery)

'==================== Removing of Older Values From the Tabel ================================


sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sCrVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
sQuery=sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

Response.Write sQuery& "<BR>"
con.execute(sQuery)

for each EntryNode in oNodDeatils.childNodes

	Response.Write "Inside "
	sEntryno=EntryNode.Attributes.Item(0).nodeValue
	dBasicAmount=EntryNode.Attributes.Item(7).nodeValue
	sItemDesc=replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
	dQty=EntryNode.Attributes.Item(3).nodeValue
	sUOM=EntryNode.Attributes.Item(4).nodeValue
	sAmount=EntryNode.Attributes.Item(2).nodeValue
	dRate=EntryNode.Attributes.Item(6).nodeValue
	dDisPer=EntryNode.Attributes.Item(8).nodeValue
	dDisAmount=EntryNode.Attributes.Item(9).nodeValue
	dItemCode=EntryNode.Attributes.Item(10).nodeValue
	dClassCode=EntryNode.Attributes.Item(11).nodeValue
	IF Trim(CStr(dItemCode)) = "" Or Not IsNumeric(dItemCode) Then
		dItemCode = 0
	End IF
	IF Trim(CStr(dClassCode)) = "" Or Not IsNumeric(dClassCode) Then
		dClassCode = 0
	End IF

	sEntryType="D"
	sNarration=""

	for each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
				sAccCode=HeaderNode.Attributes.Item(0).nodeValue
				sAccType=HeaderNode.Attributes.Item(4).nodeValue
		end if 'End of Check for Account head Node
	next 'End of Entry Node Loop

		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
		sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,ItemCode,ClassificationCode ) values ("
		sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
		sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","&dItemCode&","&dClassCode&" )"
Response.Write "<p>"&sQuery& "<BR>"
con.execute(sQuery)

'--------------Other details Pending---------------------
next'End of Voucher Node Loop


'Response.End
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
		sEntryType = "D"
	else
		sEntryType = "C"
		'dTaxAmount=dTaxAmount*-1
		dTaxAmount = Abs(dTaxAmount)
		dTaxPer = abs(dTaxPer)
	End if
	'if CInt(sAccCode)>0 then
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
	'end if
Next

dim iAdvTranNo,dAdvAmount,iCrAdvNo,sAdjNarr,sAdjTy,iCrRecNo
dAdjAmtTotal=0
IF IsObject(oNodAdvRoot) Then
	for each EntryNode in oNodAdvRoot.childNodes
		iAdvTranNo=EntryNode.Attributes.Item(0).nodeValue
		dAdvAmount=EntryNode.Attributes.Item(5).nodeValue
		iCrAdvNo = EntryNode.Attributes.Item(7).nodeValue
		sAdjTy = EntryNode.Attributes.Item(8).nodeValue

		if CDbl(dAdvAmount)>0 then
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)

	'----------- UPDATION FOR THE ADVANCE ADJUSTMENT DETAILS ----------------------
	'----------- INCLUDED ON 04/05/2004 ----------------------

			IF CStr(sAdjTy) = "I" Then
				sQuery = "Update Acc_T_CreatedAdvances set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
						 " where CreatedAdvanceNo = "&iCrAdvNo

				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)

				sAdjNarr = "Adjusted to Purchase Invoice No "&sCrVouNo&" DT: "& sVoucDate &" Amt: "& dTotal
				sQuery = "INSERT INTO ACC_T_CreatedAdvanceAdj (CreatedAdvanceNo, CreatedTransNo,  "&_
						 "AdjustedOn, AmountAdjusted, Narration)  "&_
						 "VALUES ("&iCrAdvno&", "&iCrTransNo&", Convert(Datetime,'"&sVoucDate&"',103), "&dAdvAmount&", '"&sAdjNarr&"') "

				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)
			Else
				Response.Clear
				sQuery = "Select DRCreatedReceivable From Acc_T_Receivables Where ReceivableNumber = "&iAdvTranNo&" "

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
						 "getdate(),"&dAdvAmount&")"

				Response.Write sQuery &"<br><br>"
				Con.Execute sQuery

			End IF
	'------------ END OF UPDATION ---------------------
		end if

	Next
End IF
dim sPayNarration,iCrPayableNo
if CDbl(dTotal)>CDbl(dAdjAmtTotal) then
	sQuery="select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
	objRs.open sQuery,con
		iCrPayableNo=objRs(0)
	objRs.Close
	sPayNarration="PUR INV No:"&sPurInvNo &" Dt:"&sPurInvDt

	sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
			"VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
			" PartyBillDate, AmountPayable, AmountPaid,Narration)values("&iCrPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
			"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0,'"&sPayNarration&"')"

	Response.Write sQuery &"<br><br>"
	con.execute(sQuery)

'----------- UPDATION FOR THE ADVANCE ADJUSTMENT DETAILS ----------------------
'----------- INCLUDED ON 04/05/2004 ----------------------

	sQuery="insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
				"PaidOn,AmountPaid) Values ("&iCrPayableNo&","&iCrTransNo&","&_
				"convert(datetime,'"&sVoucDate&"',103),"&dAdjAmtTotal&")"

	Response.Write sQuery &"<br><br>"
	con.execute(sQuery)

'------------ END OF UPDATION ---------------------

	sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
			 "Values("&iCrTransNo&",1,"&sAppBy&" ) "
Response.Write sQuery
	Con.Execute sQuery
end if


'Con.RollbackTrans
'Response.End
Con.CommitTrans

'============================================================================================

'******************************** Cost and ANal Procedure Call ************************

dim sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr,TempNode

sExp = "//Entry"
Set HeaderNode = oNodRoot.selectNodes(sExp)
IF HeaderNode.length <> 0 Then

	for iCtr = 0 To HeaderNode.length - 1

		sEntryno=HeaderNode.Item(iCtr).Attributes.getNamedItem("No").value
		'sEntryType=HeaderNode.Item(iCtr).Attributes.getNamedItem("CRDR").value
		'sAmount=HeaderNode.Item(iCtr).Attributes.getNamedItem("Amount").value
		'sAccUnit=HeaderNode.Item(iCtr).Attributes.getNamedItem("AccUnit").value

		sExp="//Entry[@No='"&sEntryno&"']/AccHead"
		set tempNode=oNodRoot.selectNodes(sExp)
		sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value

		set nodANL=oDOM.createElement("Root")
		nodANL.setAttribute "TransNo",iCrTransNo
		nodANL.setAttribute "EntryNo",sEntryno
		nodANL.setAttribute "UnitCode", sorgid
		nodANL.setAttribute "GlHead",sAccCode
		nodANL.setAttribute "ACTFlag","C"



		sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
		set tempNode=oNodRoot.selectNodes(sExp)
		if tempNode.length >0 then
			set EntryNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(EntryNode)
			bAddFlag=true
		end if

		sExp="//Entry[@No='"&sEntryno&"']/Analytical"
		set tempNode=oNodRoot.selectNodes(sExp)
		if tempNode.length >0 then
			set EntryNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(EntryNode)
			bAddFlag=true
		end if
		if bAddFlag then
		  Dim adoConn

		   Set adoConn = Server.CreateObject("ADODB.Connection")
		   adoConn.ConnectionString = con
		   adoConn.CursorLocation = 3
		   adoConn.Open

		   sQuery = "Proc_VouCCANALUpdate"


		   Dim adoCmd
		   Set adoCmd = Server.CreateObject("ADODB.Command")
		   Set adoCmd.ActiveConnection =adoConn
		   adoCmd.CommandText = sQuery
		   adoCmd.CommandType = 4 'adCmdStoredProc
		   adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(nodANL.xml),nodANL.xml)



		   Dim adoRS
		   Set adoRS = adoCmd.Execute()

		end if

	next 'End of Entry Node Loop
End IF

'******************************** Coast and ANal Procedure Call ************************


'----To code for Tax Form deatils----------------

if con.Errors.count <>0 then
	con.RollbackTrans

	for iCounter=0 to con.Errors.count-1
		Response.Write con.Errors(iCounter).Description &"<br>"
		Response.Write con.Errors(iCounter).Source &"<br>"
	next
	'Redirect to Error Handling System
else

	'con.CommitTrans

	'oDOM.load server.MapPath("../temp/transaction/Voucher Entry_PUR_"&Session.SessionID&".xml")

	Set newElem  = oDOM.createAttribute("CreatedTransNo")
	newElem.value = iCrTransNo
	oNodRoot.setAttributeNode(newElem)

	Set newElem  = oDOM.createAttribute("CreatedVouNo")
	newElem.value = sCrVouNo
	oNodRoot.setAttributeNode(newElem)

'	Set newElem  = oDOM.createAttribute("TransNo")
'	newElem.value = iTransNo
'	oNodRoot.setAttributeNode(newElem)

'	Set newElem  = oDOM.createAttribute("VouNo")
'	newElem.value = iVouNo
'	oNodRoot.setAttributeNode(newElem)

	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")
	objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml"))
	Response.Redirect ("VouPURDisplay.asp?TransNo="&iCrTransNo&"&CallFrm=A"&"&hFormVal="&sFormVal)
	'Server.Transfer("VouPURDisplay.asp?TransNo="&iCrTransNo)
end if

%>
