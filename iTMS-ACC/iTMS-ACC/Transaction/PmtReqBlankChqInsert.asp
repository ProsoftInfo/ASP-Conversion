<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtReqBlankChqInsert
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February, 2003
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
<%
dim objRs,sQuery,iPymtNo
dim sOrgId,iAccCode,sPayTo,dAmount,sReason,sPymtStatus,sTemp
dim sPartyType,sPartySubType,sAccType,iUserId,nReqFrom

sOrgId=Request("hUnitId")
iAccCode=Request("hAccountCode")
sPayTo=Request("txtPayTo")
dAmount=Request("txtAmount")
sReason=Request("txtReason")
sAccType=Request("selAcctype")
iUserId=Request("selUserId")

if trim(dAmount)="" then dAmount=0
sPymtStatus="010201" 'Crearted For Approval
nReqFrom = "1"	'Indicates Entry From Accounts

set objRs  = server.CreateObject("adodb.recordset")

con.BeginTrans

sQuery="select isnull(max(PaymentRequestNo),0)+1 from Acc_T_PaymentRequestHdr"
objRs.open sQuery,con
	iPymtNo=objRs(0)
objRs.Close

'For PaymentRequestCode field Now insert PaymentReqNo.Later NoSeries ill be generate n inserted.

sQuery="INSERT INTO Acc_T_PaymentRequestHdr(PaymentRequestNo, PaymentRequestDt,PaymentRequestCode, PaymentMode, RequestFor,"&_
		" RequestedBy, ApprovedBy, ApprovedOn,RequestedFromAppl) VALUES("&_
		""&iPymtNo&",getdate(),'"&iPymtNo&"','B','B',"&getUserid&",NULL,NULL,"& nReqFrom &")"
con.execute(sQuery)

'Response.Write sAccType
if sAccType="S" then
	sQuery="INSERT INTO Acc_T_PaymentRequestDet(PaymentRequestNo, RequestEntryNo, TransactionNumber, PayablesNumber, AccountingUnit,"&_
		" AccUnitAccountHead, AccUnitPartyType, AccUnitPartySubType, AccUnitPartyCode, ToBePaidTo, ReasonForPayment, AmountToPay,"&_
		" ToPayBefore, PrincipalAmountToPay, InterestAmountToPay, AmountPaid, PaidOn, PrincipalAmountPaid, InterestAmountPaid, StatusOfRequestDet)"&_
		"VALUES("&iPymtNo&", 1,NULL,NULL,'"&sOrgId&"',NULL,NULL,NULL,NULL,'"&sPayTo&"','"&sReason&"',"&dAmount&",NULL,NULL,NULL,NULL,NULL,NULL,NULL,'"&sPymtStatus&"')"
			
elseif sAccType="G" then
	sQuery="INSERT INTO Acc_T_PaymentRequestDet(PaymentRequestNo, RequestEntryNo, TransactionNumber, PayablesNumber, AccountingUnit,"&_
		" AccUnitAccountHead, AccUnitPartyType, AccUnitPartySubType, AccUnitPartyCode, ToBePaidTo, ReasonForPayment, AmountToPay,"&_
		" ToPayBefore, PrincipalAmountToPay, InterestAmountToPay, AmountPaid, PaidOn, PrincipalAmountPaid, InterestAmountPaid, StatusOfRequestDet)"&_
		"VALUES("&iPymtNo&", 1,NULL,NULL,'"&sOrgId&"',"&iAccCode&",NULL,NULL,NULL,'"&sPayTo&"','"&sReason&"',"&dAmount&",NULL,NULL,NULL,NULL,NULL,NULL,NULL,'"&sPymtStatus&"')"
else
	sTemp=Split(iAccCode,"?")
	sPartyType=sTemp(0)
	sPartySubType=sTemp(1)
	iAccCode=sTemp(3)
	sQuery="INSERT INTO Acc_T_PaymentRequestDet(PaymentRequestNo, RequestEntryNo, TransactionNumber, PayablesNumber, AccountingUnit,"&_
		" AccUnitAccountHead, AccUnitPartyType, AccUnitPartySubType, AccUnitPartyCode, ToBePaidTo, ReasonForPayment, AmountToPay,"&_
		" ToPayBefore, PrincipalAmountToPay, InterestAmountToPay, AmountPaid, PaidOn, PrincipalAmountPaid, InterestAmountPaid, StatusOfRequestDet)"&_
		"VALUES("&iPymtNo&", 1,NULL,NULL,'"&sOrgId&"',NULL,'"&sPartyType&"',"&sPartySubType&","&iAccCode&",'"&sPayTo&"','"&sReason&"',"&dAmount&",NULL,NULL,NULL,NULL,NULL,NULL,NULL,'"&sPymtStatus&"')"
		
end if		
con.execute(sQuery)


sQuery="insert into Acc_T_VouchersForApproval(PaymentRequestNo,ApprovalLevel,ToBeApprovedBy)"&_
		"Values("&iPymtNo&",1,"&iUserId&")"
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

