<%@ Language=VBScript %>
<%option explicit	%>
<%'on error resume next%>
<%
	'Program Name				:	PmtGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February  11, 2003
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

Dim oDOM,Root,objRs,sQuery,sPayTo
dim EntryNode,RequestNode,nodANL,newElem,TempNode
dim sAccHeadCode,sPartySubType

dim sNarration,sAccount,sAddtional,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName

dim sVouCode,sVoucDate,sAccUnit

dim dTotal,sTransType
dim sAccType,sAccCode,iTransNo
dim sDocType,sVouStatus
dim iPmtNo,iCounter,iEntryNo
dim iSeriesNo,iSeriesCode
dim dAddTotal,iAdvNo,Insnode
dim sBankDet,sBankInsType,sBankInsNo,sPayAt,DrwOnBank,sBankTemp
dim dBankInsDate

sVouStatus="010101" 'Crearted For Accounting to be Approved

sVouCode=Request("hVouCode")
sVouName=Request("hVouName")
iPmtNo=Request("hPaymentNo")
sBankDet = Request("hInsDet")

If sBankDet <> "" then
	sBankTemp = split(sBankDet,":")
	sBankInsType = sBankTemp(2)
	sBankInsNo   = sBankTemp(1)
	dBankInsDate = FormatDate(sBankTemp(5))
	sPayAt		 = sBankTemp(3)
	DrwOnBank    = sBankTemp(4)
	
	'sBankInsType = sBankTemp(0)
	'sBankInsNo   = sBankTemp(1)
	'dBankInsDate = FormatDate(sBankTemp(2))
	'sPayAt		 = sBankTemp(3)
	'DrwOnBank    = sBankTemp(4)
Else		
	sBankInsType = "NULL"
	sBankInsNo   = "NULL"
	dBankInsDate = "01/01/1900"
	sPayAt		 = "NULL"
	DrwOnBank    = "NULL"
End IF

set objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.Load server.MapPath("../temp/transaction/Payment Request_"&sVouName&"_"&Session.SessionID&".xml")	

set Root=oDOM.documentElement
sOrgId=Root.Attributes.Item(0).nodeValue
sBookNo =Root.Attributes.Item(2).nodeValue
sVouType=Root.Attributes.Item(4).nodeValue
sVoucDate=Root.Attributes.Item(5).nodeValue
'sAccHeadCode=Root.Attributes.Item(6).nodeValue
sAccHeadCode = Request("selBookId")
sTemp = Split(sAccHeadCode,"?")
sAccHeadCode = sTemp(1)

for each EntryNode in Root.childNodes
	if EntryNode.nodeName="Entry" then 
		dTotal=CDbl(dTotal)+CDbl( EntryNode.Attributes.Item(3).nodeValue)
		sPayTo = EntryNode.Attributes.Item(2).nodeValue
	end if
next	


sTransType=sVouName&"P"


sQuery="select CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
	"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo&" "
' Response.Write sQuery
objRs.open sQuery,con
    if not objrs.eof then
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
		
sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,CrDrIndication,"
sQuery=sQuery&"BankInstrumentType,BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,"
sQuery=sQuery&"CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,PayToRecdFrom) values"
sQuery=sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',NULL,"
sQuery=sQuery&""&sAccHeadCode&",'"&iVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&","
sQuery=sQuery&"'"&sVouType&"','"&sBankInsType&"','"&sBankInsNo&"',convert(datetime,'"&dBankInsDate&"',103),"
sQuery=sQuery&"'"&sPayAt&"','"&DrwOnBank&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"', '"&sPayTo&"')"
	
Response.Write sQuery &"<br><br>"

con.execute(sQuery)

for each EntryNode in Root.childNodes
	if EntryNode.nodeName="RequestDetails" then 
		set RequestNode=EntryNode
	elseif EntryNode.nodeName="Entry" then 
		iEntryNo=EntryNode.Attributes.Item(0).nodeValue 
		sAmount=EntryNode.Attributes.Item(3).nodeValue 
		sAccUnit=EntryNode.Attributes.Item(4).nodeValue 
		
		for each TempNode in EntryNode.childNodes
		
			if TempNode.nodeName="AccHead" then
				sAccCode=TempNode.Attributes.Item(0).nodeValue 
				sAccType=TempNode.Attributes.Item(4).nodeValue
				
								
			end if 'End of Check for Account head Node
			if 	TempNode.nodeName="Narration" then
					sNarration=TempNode.text
			end if 'End of Check for Narration Node
		next 'End of Entry Node Loop
		dAddTotal=0	
			 
		'Response.Write "<p>sAccType = "&sAccType
			 
		if StrComp(sAccType,"G")=0 then
			sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"&_
					" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"&_
					" VoucherNarration, Amount,TransCrDrIndication) values ("&_
					""&iTransNo&",'"& sAccUnit & "'" &_
					","&iEntryNo&","&sAccCode&",NULL,NULL,NULL,"&_
					" '"&sNarration&"',"&sAmount&",'D')"
		else
		
			sTemp=Split(sAccCode,"?")
			'Response.Write "<p>sAccCode="&sAccCode & "<p>"
			If sTemp(1) <> "" Then
				sPartySubType = sTemp(1)
			Else
				sPartySubType = "NULL"
			End IF
			sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"&_
				" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,"&_
				" VoucherNarration, Amount,TransCrDrIndication) values ("&_
				""& iTransNo&",'"&sAccUnit&"'" &_
				","&iEntryNo&",NULL,'"&sTemp(0)&"',"&sPartySubType&","&sTemp(3)&","&_
				" '"&sNarration&"',"&sAmount&",'D')"
		end if
		Response.Write sQuery& "<BR>"	
		con.execute(sQuery)
		
		
	end if	
next'End of Voucher Node Loop

IF sAccType="P" and CDbl(sAmount)> CDbl(dAddTotal) THEN
	sQuery = "Select isNull(Max(CreatedAdvanceNo),0)+1 from Acc_T_CreatedAdvances "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		iAdvNo = objRs(0)
	End IF
	objRs.Close
	
	sQuery="INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
		"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
		" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sPartySubType&","&_
		""&sTemp(3)&","&sAmount&","&CDbl(sAmount)- CDbl(dAddTotal)&",NULL,NULL)"
	Response.Write sQuery	&"<br><br>"
	con.execute(sQuery)	
END IF	

sQuery = "Update Acc_T_PaymentRequestHdr Set CreatedTransNo = "& iTransNo&" Where PaymentRequestNo = "&iPmtNo&" "
Response.Write "<p>sQuery = "& sQuery
con.Execute sQuery

sQuery = "Update Acc_T_PaymentRequestDet Set AmountPaid = isNull(AmountPaid,0) + "&dTotal&" Where  "&_
		 "PaymentRequestNo = "&iPmtNo&" and RequestEntryNo = "&iEntryNo
Response.Write sQuery	&"<br><br>"
Con.Execute sQuery
	
'Added By UmaMaheswari S, On April 06,2011
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source ="Select Distinct RequestEntryNo From Acc_T_PaymentRequestDet Where PaymentRequestNo="&iPmtNo&" "
	.ActiveConnection = con
	.Open 
End With
Do While Not objRs.EOF 
	sQuery="update Acc_T_PaymentRequestDet set StatusOfRequestDet='010204' where PaymentRequestNo="&iPmtNo&" and RequestEntryNo="&objrs(0)
	Response.Write sQuery& "<BR>"	
	con.execute(sQuery)		
	
	objRs.MoveNext
Loop
objRs.Close 


for each TempNode in RequestNode.childNodes		
	iEntryNo=TempNode.Attributes.Item(0).nodeValue
	sQuery="update Acc_T_PaymentRequestDet set StatusOfRequestDet='010204' where PaymentRequestNo="&iPmtNo&" and RequestEntryNo="&iEntryNo
	Response.Write sQuery& "<BR>"	
	con.execute(sQuery)		
next
	

'Response.End 	
if con.Errors.count <>0 then
	con.RollbackTrans
	
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter)
	next
	'Redirect to Error Handling System
else
	con.CommitTrans
	'con.RollbackTrans
	'Response.End 

	'Set newElem  = oDOM.createAttribute("TransNo")
	'newElem.value = iTransNo
	'Root.setAttributeNode(newElem)
	
	'Set newElem  = oDOM.createAttribute("VoucherNo")
	'newElem.value = iVouNo
	'Root.setAttributeNode(newElem)
	
	'Set newElem  = oDOM.createAttribute("RequestNo")
	'newElem.value = iPmtNo
	'RequestNode.setAttributeNode(newElem)
	
	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")	
	
	'Response.Redirect ("PmtBADisplay.asp?TransNo="&iTransNo &"&VouName="&sVouName)
	Response.Redirect ("PAYMENTREQUESTS.ASP")
end if
%>

	
