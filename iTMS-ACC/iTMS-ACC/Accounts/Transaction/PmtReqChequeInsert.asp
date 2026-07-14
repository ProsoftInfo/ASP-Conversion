<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtReqChequeInsert.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 21, 2003
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
<%
dim objRs,sQuery,iPymtNo
dim sOrgId,iAccCode,sPayTo,dAmount,sReason,sPymtStatus,sTemp
dim sPartyType,sPartySubType,sAccType,iUserId
dim oDOM,EntryNode,Root,sReqType,sReqFor
dim iSno,iTransNo,iPayableNo,nReqFrom

sOrgId=Request("hUnitId")
iAccCode=Request("hAccountCode")
sPayTo=Request("txtPayTo")
sAccType=Request("selAcctype")
iUserId=Request("selUserId")
sReqType=Request("hRequestType")

if sReqType="C" then
	sReqFor="O"
else
	sReqFor="A"
end if

sPymtStatus="010201" 'Crearted For Approval
nReqFrom = "1"	'Indicates Entry From Accounts

set objRs  = server.CreateObject("adodb.recordset")
' Create our DOM Document Objects
SET oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.Load server.MapPath("../temp/transaction/Payment Requestion_CHQ_"&Session.SessionID&".xml")


con.BeginTrans

sQuery="select isnull(max(PaymentRequestNo),0)+1 from Acc_T_PaymentRequestHdr"
objRs.open sQuery,con
	iPymtNo=objRs(0)
objRs.Close

'For PaymentRequestCode field Now insert PaymentReqNo.Later NoSeries ill be generate n inserted.

sQuery="INSERT INTO Acc_T_PaymentRequestHdr(PaymentRequestNo, PaymentRequestDt,PaymentRequestCode, PaymentMode, RequestFor,"&_
		" RequestedBy, ApprovedBy, ApprovedOn,RequestedFromAppl) VALUES("&_
		" "&iPymtNo&",getdate(),'"&iPymtNo&"','"&sReqType&"','"&sReqFor&"',"&getUserid&",NULL,NULL,"& nReqFrom &")"
con.execute(sQuery)

'Response.Write sQuery &"<br>"

sTemp=Split(iAccCode,"?")
sPartyType=sTemp(0)
sPartySubType=sTemp(1)
iAccCode=sTemp(3)

iSno=1

SET Root=oDOM.documentElement
FOR EACH EntryNode IN Root.childNodes
	IF CDbl(EntryNode.Attributes.Item(6).nodeValue)>0 THEN
		iTransNo=EntryNode.Attributes.Item(0).nodeValue
		iPayableNo=EntryNode.Attributes.Item(1).nodeValue
		sReason="Payment For Invoice:"&EntryNode.Attributes.Item(3).nodeValue &"-"&EntryNode.Attributes.Item(4).nodeValue
		dAmount=CDbl(EntryNode.Attributes.Item(6).nodeValue)

		sQuery="INSERT INTO Acc_T_PaymentRequestDet(PaymentRequestNo, RequestEntryNo, TransactionNumber, PayablesNumber, AccountingUnit,"&_
			" AccUnitAccountHead, AccUnitPartyType, AccUnitPartySubType, AccUnitPartyCode, ToBePaidTo, ReasonForPayment, AmountToPay,"&_
			" ToPayBefore, PrincipalAmountToPay, InterestAmountToPay, AmountPaid, PaidOn, PrincipalAmountPaid, InterestAmountPaid, StatusOfRequestDet)"&_
			"VALUES("&iPymtNo&", "&iSno&",NULL,NULL,'"&sOrgId&"',NULL,'"&sPartyType&"',"&sPartySubType&","&iAccCode&",'"&sPayTo&"','"&sReason&"',"&dAmount&",NULL,NULL,NULL,NULL,NULL,NULL,NULL,'"&sPymtStatus&"')"
		con.execute(sQuery)
		Response.Write sQuery &"<br>"
		iSno=CInt(iSno)+1
	END IF
NEXT


sQuery="insert into Acc_T_VouchersForApproval(PaymentRequestNo,ApprovalLevel,ToBeApprovedBy)"&_
		"Values("&iPymtNo&",1,"&iUserId&")"
'Response.Write sQuery &"<br>"

con.execute(sQuery)

if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	con.CommitTrans
	'Response.Redirect ("PmtReqTypeSelection.asp")
	Response.Redirect ("PAYMENTREQUESTS.ASP")
end if
%>

