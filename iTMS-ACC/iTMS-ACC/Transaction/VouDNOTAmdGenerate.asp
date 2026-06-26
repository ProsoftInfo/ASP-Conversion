<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNAmdGenerate.asp
	'Module Name				:	ACCOUNTS (Debit Note Other Voucher Amendment Updation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Aug 04, 2004
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

Dim oDOM,nodHeader,Root,objRs,sQuery
dim EntryNode,HeaderNode,nodANL,newElem

dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus,sParType,sParSubType,sParCode,iPayableNo
dim iSeriesNo,iSeriesCode,sPayTo,bAddFlag,sVouTransNo,iAmdendmentNo,objfs

dim sAccHeadCode

sApprove=Request("optApprove")
sVouCode=Request("hVouCode")
sVouName=Request("hVouName")
sVouCode = "06"
sVouName = "DN"

sVouStatus="010101" 'Crearted For Approval

set objRs  = server.CreateObject("adodb.recordset")
Set objfs = CreateObject("Scripting.FileSystemObject")


' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.Load server.MapPath("../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml")




set Root=oDOM.documentElement

sOrgId=Root.Attributes.getNamedItem("UnitNo").value

sBookNo =Root.Attributes.getNamedItem("BookNo").value

Response.Write sBookNo &" ================ "&"<br><br>"
sVouType=Root.Attributes.getNamedItem("CRDR").value
sVoucDate=Root.Attributes.getNamedItem("VouDate").value
iTransNo = Root.Attributes.getNamedItem("TransNo").value
sVouTransNo = Root.Attributes.getNamedItem("VoucherNo").value

sTemp=Split(Root.Attributes.getNamedItem("PartyCode").value,"?")

sParType=trim(sTemp(0))
sParSubType=trim(sTemp(1))
sParCode=trim(sTemp(3))

Root.Attributes.getNamedItem("Approver").value=sApprove
sPayTo=Replace(Root.childNodes(0).Attributes.getNamedItem("Payto").value,"'","''")

FOR EACH EntryNode IN Root.childNodes
		dTotal=dTotal+CDbl(EntryNode.Attributes.Item(3).nodeValue)
NEXT

sTransType=sVouName&"R"

Con.BeginTrans

sQuery = "Select CreatedVouchStatus From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouStatus = objRs(0)
End IF
objRs.Close

'--------------Insert into Histroy tables---------------------------
sQuery="select isnull(max(AmendmentNo),0)+1 from Acc_T_HistoryVoucherHeader where "&_
		"TransactionNumber="&iTransNo
		'Response.Write sQuery
objRs.open sQuery,con
if not objRs.EOF then
	iAmdendmentNo=objRs(0)
end if
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
	"'A',"&getUserid&",getdate(),'Amendment' "&_
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

sQuery = "INSERT INTO Acc_T_HisCreatedPayables (AmendmentNo, PayablesNumber, CreatedTransNo, "&_
		 "OUDefinitionID, PartyType, PartySubType, PartyCode, VoucherDate, PartyBillNumber, "&_
		 "PartyBillDate, AmountPayable, AmountPaid, Narration) "&_
		 "SELECT "&iAmdendmentNo&", PayablesNumber, CreatedTransNo, OUDefinitionID, PartyType, PartySubType, PartyCode, "&_
		 "VoucherDate, PartyBillNumber, PartyBillDate,AmountPayable, AmountPaid, Narration "&_
		 "FROM Acc_T_CreatedPayables Where CreatedTransNo = "&iTransNo

con.execute(sQuery)

sQuery = "INSERT INTO Acc_T_HisCreateReceivables (AmendmentNo, ReceivableNumber, CreatedTransNo, "&_
		 "OUDefinitionID, VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber, "&_
		 "PartyInvoiceDate, AmountReceivable, AmountReceived, Narration) "&_
		 "SELECT "&iAmdendmentNo&" , ReceivableNumber, CreatedTransNo, OUDefinitionID, VoucherDate, PartyType, "&_
		 "PartySubType, PartyCode, PartyInvoiceNumber, PartyInvoiceDate, AmountReceivable, "&_
		 "AmountReceived, Narration FROM Acc_T_CreatedReceivables "&_
		 "Where CreatedTransNo = "&iTransNo

con.execute(sQuery)

'----------Delete Reocords-----------------------------------------
sQuery="delete Acc_T_CretedVoucherAHDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherCCDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedPayables where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedReceivables where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedAdvances where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
con.execute(sQuery)


'============== Deletion of Records Ends ===========================================


sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
		 "PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
		 "CrDrIndication,BankInstrumentType,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"&_
		 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
		 "'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sVouTransNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
		 ",'"&sVouType&"','OT',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

Response.Write sQuery& "<BR><BR>"
con.execute(sQuery)
'-----------------------PROCESS ENTRY NODES-------------------------------
FOR EACH EntryNode IN Root.childNodes
	sEntryno=EntryNode.Attributes.Item(0).nodeValue
	sAmount=EntryNode.Attributes.Item(3).nodeValue
	sEntryType=EntryNode.Attributes.Item(1).nodeValue
	sAccUnit=EntryNode.Attributes.Item(4).nodeValue
'---------PROCESS THE CHILD NODES OF ENTRIES FOR DETAIL TABLE UPDATION----
	FOR EACH HeaderNode IN EntryNode.childNodes
		IF HeaderNode.nodeName="AccHead" THEN
				sAccCode=HeaderNode.Attributes.Item(0).nodeValue
				sAccType=HeaderNode.Attributes.Item(4).nodeValue
		END IF 'End of Check for Account head Node
		IF 	HeaderNode.nodeName="Narration" THEN
				sNarration=HeaderNode.text
		END IF 'End of Check for Narration Node
	NEXT
'-------------END OF PROCESSING CHILD NODES OF ENTRIES---------------------
'----------------------------DETAIL TABLE UPDATION-------------------------
	IF StrComp(sAccType,"G")=0 THEN
		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication) values ("
		sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
		sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"')"
	ELSE
		sTemp=Split(sAccCode,"?")

		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication) values ("
		sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
		sQuery=sQuery& ","&sEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(1)&","&sTemp(3)&","
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"')"
	END IF
Response.Write sQuery& "<BR><BR>"
con.execute(sQuery)
Next
'-----------------------END OF DETAIL TABLE UPDATION----------------------
if sVouName="CN" then

	sQuery="select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
	objRs.open sQuery,con
	if not objRs.EOF then
		iPayableNo=objRs(0)
	end if
	objRs.Close

sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
		"VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
		" PartyBillDate, AmountPayable, AmountPaid,Narration)values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
		"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",NULL,"&_
		"NULL,"&dTotal&",0,'"&sPayTo&"')"

	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)
else
	sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_CreatedReceivables"
	objRs.open sQuery,con
	if not objRs.EOF then
		iPayableNo=objRs(0)
	end if
	objRs.Close

	sQuery="INSERT INTO Acc_T_CreatedReceivables(ReceivableNumber, CreatedTransNo, OUDefinitionID,"&_
			"VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber,"&_
			" PartyInvoiceDate, AmountReceivable, AmountReceived,Narration)values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",NULL,"&_
			"NULL,"&dTotal&",0,'"&sPayTo&"')"

	Response.Write sQuery& "<BR><BR>"
	con.execute(sQuery)
end if


'con.RollbackTrans
'Response.End 
con.CommitTrans





'******************************** Cost and ANal Procedure Call ************************

dim sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr,TempNode,sExp

sExp = "//Entry"
Set HeaderNode = Root.selectNodes(sExp)
IF HeaderNode.length <> 0 Then

	for iCtr = 0 To HeaderNode.length - 1

		sEntryno=HeaderNode.Item(iCtr).Attributes.getNamedItem("No").value

		sExp="//Entry[@No='"&sEntryno&"']/AccHead"
		set tempNode=Root.selectNodes(sExp)
		sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value

		set nodANL=oDOM.createElement("Root")
		nodANL.setAttribute "TransNo",iTransNo
		nodANL.setAttribute "EntryNo",sEntryno
		nodANL.setAttribute "UnitCode", sorgid
		nodANL.setAttribute "GlHead",sAccCode
		nodANL.setAttribute "ACTFlag","C"

		sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
		set tempNode=Root.selectNodes(sExp)
		if tempNode.length >0 then
			set EntryNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(EntryNode)
			bAddFlag=true
		end if

		sExp="//Entry[@No='"&sEntryno&"']/Analytical"
		set tempNode=Root.selectNodes(sExp)
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
		   Dim adoCmd,adoRS
		   Set adoCmd = Server.CreateObject("ADODB.Command")
		   Set adoCmd.ActiveConnection =adoConn
		   adoCmd.CommandText = sQuery
		   adoCmd.CommandType = 4 'adCmdStoredProc
		   adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(nodANL.xml),nodANL.xml)
		   Set adoRS = adoCmd.Execute()
		end if
	next 'End of Entry Node Loop
End IF

'******************************** Coast and ANal Procedure Call ************************

if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
else

	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
	if objfs.FileExists(server.MapPath("../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml")) then
	    objfs.DeleteFile(server.MapPath("../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml"))
	end if'if objfs.FileExists(server.MapPath("../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml")) then
	Response.Redirect ("VouDNOtherDisplay.asp?TransNo="&iTransNo)
end if

%>

