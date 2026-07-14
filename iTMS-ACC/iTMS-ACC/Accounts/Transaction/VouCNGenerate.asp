<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 16 , 2003
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

Dim oDOM,nodHeader,Root,objRs,sQuery
dim EntryNode,HeaderNode,nodANL,newElem

dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus,sParType,sParSubType,sParCode,iPayableNo
dim iSeriesNo,iSeriesCode,sPayTo,bAddFlag,dTdsAmt,sTdsPer,Checknode,sApprover,sSelInvNo


dim sAccHeadCode,sCrtVouType
dim sChkVal,sPayAt,sDrwOn,iBookNo

iBookNo = Request("hBookcode")
sChkVal = Request("hChkVal")

sApprove=Request("optApprove")
sVouCode=Request("hVouCode")
sVouName=Request("hVouName")
sApprover = Request.Form("selUserID")

sCrtVouType =Request("hCrtVouType")
'Response.Write "sCrtVouType="&sCrtVouType

IF CStr(Request.Form("hSelVouTy")) <> "OT" Then
	sSelInvNo = Request.Form("hInvNos")
End IF


IF Cstr(sApprove) = "Y" Then
	sVouStatus="010101" 'Crearted For Approval
Else
	sVouStatus="010103" 'Crearted For Accounting
End IF

set objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.Load server.MapPath("../temp/transaction/Voucher Entry_"&sVouName&"_"&Session.SessionID&".xml")

set Root=oDOM.documentElement

sOrgId=Root.Attributes.getNamedItem("UnitNo").value

sBookNo =Root.Attributes.getNamedItem("BookNo").value
sVouType=Root.Attributes.getNamedItem("CRDR").value
sVoucDate=Root.Attributes.getNamedItem("VouDate").value

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

sQuery="select CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
	"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo

Response.Write sQuery
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
	iTransNo=objRs(0)
objRs.Close

If trim(sChkVal) = "Y" then  
	sCrtVouType = "G" 
	sDrwOn = iBookNo
Else
	sDrwOn = "NULL"	 
End If

sQuery="insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
	"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
	"CrDrIndication,BankInstrumentType,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,PayableAt,DrawnOnBank) values"&_
	"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
	"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
	",'"&sVouType&"','OT',"&getUserid&",getdate(),NULL,'"&sVouStatus&"','"&sCrtVouType&"','"&sDrwOn&"')"


Response.Write sQuery& "<BR><BR>"
con.execute(sQuery)
'-----------------------PROCESS ENTRY NODES-------------------------------
FOR EACH EntryNode IN Root.childNodes
	sExp = "//Entry[@TdsAmount]"
	Set Checknode = Root.selectNodes(sExp)

	IF Checknode.length <> 0 Then
		dTdsAmt = EntryNode.Attributes.Item(6).nodeValue
		sTdsPer = EntryNode.Attributes.Item(8).nodeValue
	Else
		dTdsAmt = 0
		sTdsPer = 0
	End IF
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
	sNarration = Replace(sNarration,"'"," ")
	IF StrComp(sAccType,"G")=0 THEN
		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSPercentage) values ("
		sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
		sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&dTdsAmt&", "&sTdsPer&")"
	ELSE
		sTemp=Split(sAccCode,"?")

		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSPercentage) values ("
		sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
		sQuery=sQuery& ","&sEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(1)&","&sTemp(3)&","
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&dTdsAmt&", "&sTdsPer&")"
	END IF
Response.Write sQuery& "<BR><BR>"

con.execute(sQuery)
Next
'-----------------------END OF DETAIL TABLE UPDATION----------------------
if sVouName="CN" then

	sQuery="select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
	objRs.open sQuery,con
		iPayableNo=objRs(0)
	objRs.Close
	IF trim(sCrtVouType) = "G" then 
		sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
				"VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
				" PartyBillDate, AmountPayable, AmountPaid,Narration)values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",NULL,"&_
				"NULL,"&dTotal&","&dTotal&",'"&sPayTo&"')"
	Else
		sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
				"VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
				" PartyBillDate, AmountPayable, AmountPaid,Narration)values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",NULL,"&_
				"NULL,"&dTotal&",0,'"&sPayTo&"')"
	End IF
	Response.Write sQuery &"<BR><BR>"
	con.execute(sQuery)
else
	sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_CreatedReceivables"
	objRs.open sQuery,con
		iPayableNo=objRs(0)
	objRs.Close
	IF trim(sCrtVouType) = "G" then 
		sQuery="INSERT INTO Acc_T_CreatedReceivables(ReceivableNumber, CreatedTransNo, OUDefinitionID,"&_
				"VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber,"&_
				" PartyInvoiceDate, AmountReceivable, AmountReceived,Narration)values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",NULL,"&_
				"NULL,"&dTotal&","&dTotal&",'"&sPayTo&"')"
	Else
		sQuery="INSERT INTO Acc_T_CreatedReceivables(ReceivableNumber, CreatedTransNo, OUDefinitionID,"&_
				"VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber,"&_
				" PartyInvoiceDate, AmountReceivable, AmountReceived,Narration)values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",NULL,"&_
				"NULL,"&dTotal&",0,'"&sPayTo&"')"
	End IF
	Response.Write sQuery &"<BR><BR>"
	con.execute(sQuery)
end if

IF CStr(sApprover) = "" Then
	sApprover = getUserID()
End IF

IF CStr(sApprove) = "Y" Then

	sQuery = "insert into Acc_T_VouchersForApproval(CreatedTransNo,ApprovalLevel,ToBeApprovedBy)"&_
			 "Values("&iTransNo&",1,"&sApprover&")"
	con.execute(sQuery)
End IF

IF Len(sSelInvNo) <> 0 Then
	sTemp = Split(sSelInvNo,",")
	For iCtr = 0 To UBound(sTemp)
		IF CStr(Trim(sTemp(iCtr))) <> "" Then
			sQuery = "UPDATE Acc_T_VoucherHeader SET BankInstrumentNo = '1' WHERE CreatedTransNo = "&Trim(sTemp(iCtr))&" "
			'Response.Write sQuery
			Con.Execute sQuery
		End IF
	Next
End IF
'Response.End 
'con.CommitTrans

'******************************** Cost and ANal Procedure Call ************************

dim sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr,TempNode,sExp



'******************************** Coast and ANal Procedure Call ************************

if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	con.CommitTrans

	Set newElem  = oDOM.createAttribute("TransNo")
	newElem.value = iTransNo
	Root.setAttributeNode(newElem)

	Set newElem  = oDOM.createAttribute("VoucherNo")
	newElem.value = iVouNo
	Root.setAttributeNode(newElem)

	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
	select case sVouName
		Case "CN" Response.Redirect ("VouCNOtherDisplay.asp?TransNo="&iTransNo&"&CrtType="&sCrtVouType)
		Case "DN" Response.Redirect ("VouDNOtherDisplay.asp?TransNo="&iTransNo&"&CrtType="&sCrtVouType)

	end select
end if

%>

