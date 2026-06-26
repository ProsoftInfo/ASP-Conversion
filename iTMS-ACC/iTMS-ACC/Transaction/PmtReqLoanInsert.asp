<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtReqLoanInsert.asp
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

dim sPartyType,sPartySubType,sAccType
dim sPayTerms,iCounter,iToCounter,sRequestType
dim iInstallNo,sDate,iInterstRate,nReqFrom
dim dPrincipal,dTermAmount,dTermInterst,iUserId

sOrgId=Request("hUnitId")
sAccType=Request("selAccType")
iAccCode=Request("hAccountCode")
sPayTo=Request("txtPayTo")

dAmount=Request("txtLoanAmount")
iInterstRate=Request("txtInterstRate")
sDate=Request("txtStartDate")
iInstallNo=Request("txtInstallmentno")
sPayTerms=Request("selpayTerms")
iUserId=Request("selUserId")
sRequestType=Request("hFlag")	

sReason=Request("txtReason")	

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
		""&iPymtNo&",getdate(),'"&iPymtNo&"','B','"&sRequestType&"',"&getUserid&",NULL,NULL,"& nReqFrom&")"
con.execute(sQuery)

'Response.Write "<p>iAccCode="&iAccCode

if sAccType<>"G" then
	sTemp=Split(iAccCode,"?")
	If UBound(sTemp) > 0 Then
		sPartyType=sTemp(0)
		sPartySubType=sTemp(1)
		iAccCode=sTemp(3)
	Else
		sPartyType = "CR"
		sPartySubType = "NULL"
		iAccCode = iAccCode  
	End IF
	for iCounter=1 to iInstallNo
		dTermAmount=Request("txtAmount"&iCounter)
		dPrincipal=Request("txtPrincipal"&iCounter)
		dTermInterst=Request("txtInterst"&iCounter)
		if trim (dTermAmount)<>"" then
			sQuery="INSERT INTO Acc_T_PaymentRequestDet(PaymentRequestNo, RequestEntryNo, TransactionNumber, PayablesNumber, AccountingUnit,"&_
				" AccUnitAccountHead, AccUnitPartyType, AccUnitPartySubType, AccUnitPartyCode, ToBePaidTo, ReasonForPayment, AmountToPay,"&_
				" ToPayBefore, PrincipalAmountToPay, InterestAmountToPay, AmountPaid, PaidOn, PrincipalAmountPaid, InterestAmountPaid, StatusOfRequestDet)"&_
				"VALUES("&iPymtNo&", "&iCounter&",NULL,NULL,'"&sOrgId&"',NULL,'"&sPartyType&"',"&sPartySubType&","&iAccCode&",'"&sPayTo&"','"&sReason&"',"&dTermAmount&","&_
				"convert(datetime,'"&GetInterval(sDate,sPayTerms,iCounter)&"',103),"&dPrincipal&","&dTermInterst&",NULL,NULL,NULL,NULL,'"&sPymtStatus&"')"
				
			Response.Write sQuery
			con.execute(sQuery)
		end if
	Next
else
	for iCounter=1 to iInstallNo
		dTermAmount=Request("txtAmount"&iCounter)
		dPrincipal=Request("txtPrincipal"&iCounter)
		dTermInterst=Request("txtInterst"&iCounter)
		if trim (dTermAmount)<>"" then
			sQuery="INSERT INTO Acc_T_PaymentRequestDet(PaymentRequestNo, RequestEntryNo, TransactionNumber, PayablesNumber, AccountingUnit,"&_
				" AccUnitAccountHead, AccUnitPartyType, AccUnitPartySubType, AccUnitPartyCode, ToBePaidTo, ReasonForPayment, AmountToPay,"&_
				" ToPayBefore, PrincipalAmountToPay, InterestAmountToPay, AmountPaid, PaidOn, PrincipalAmountPaid, InterestAmountPaid, StatusOfRequestDet)"&_
				"VALUES("&iPymtNo&","&iCounter&",NULL,NULL,'"&sOrgId&"',"&iAccCode&",NULL,NULL,NULL,'"&sPayTo&"','"&sReason&"',"&dTermAmount&","&_
				"convert(datetime,'"&GetInterval(sDate,sPayTerms,iCounter)&"',103),"&dPrincipal&","&dTermInterst&",NULL,NULL,NULL,NULL,'"&sPymtStatus&"')"
			con.execute(sQuery)		
			end if
	next		
end if		

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
'	Response.End 
	con.CommitTrans
	Response.Redirect ("PaymentRequests.asp")
end if
%>
<%
function GetInterval(sDate,sIntervalType,iInterval)
	dim iMonth,iYear,iDay,sDate1
	
	iDay=mid(sDate,1,2)
	iMonth=cint(mid(sDate,4,2))
	iYear=mid(sDate,7,4)
	sDate1=DateSerial(iYear,iMonth,iDay)

	select Case sIntervalType
		Case "M"
				sDate1=DateAdd("m",CInt(iInterval),DateSerial(iYear,iMonth,iDay))
		Case "Q" 
				sDate1=DateAdd("m",CInt(iInterval)*4,DateSerial(iYear,iMonth,iDay))
		Case "H" 
				sDate1=DateAdd("m",CInt(iInterval)*6,DateSerial(iYear,iMonth,iDay))
		Case "Y" 
				sDate1=DateAdd("y",CInt(iInterval),DateSerial(iYear,iMonth,iDay))
	end select
	iDay=Day(sDate1)
	iMonth=Month(sDate1)
	iYear=year(sDate1)
	if cint(iDay) <10 then
		iDay="0"&iDay
	end if
	if cint(iMonth) <10 then
		iMonth="0"&iMonth
	end if
GetInterval=cstr(iDay)&"/"&cstr(iMonth)&"/"&cstr(iYear)	
End function
%>
