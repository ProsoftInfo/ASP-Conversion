<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PaymtReqApprove.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February 10, 2003
	'Modified By				:   UmaMaheswari S, On 05th April 2011
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
dim sQuery,sReqNo,sReqStatus,sActionFlag,iAppLevel,sTemp,iCtr

sTemp=split(Request("RequestNo"),",")
iAppLevel = "1"

'if sActionFlag="A" then
	sReqStatus="010203" 'Approved
'else
'	sReqStatus="010205" 'Rejected
'end if
'Response.Write "<p>sTemp="&Request("RequestNo")
con.BeginTrans
For iCtr = LBound(sTemp) to UBound(sTemp)

	sReqNo = sTemp(iCtr)
		
	sQuery="Update Acc_T_PaymentRequestDet set StatusOfRequestDet='"&sReqStatus&"' where PaymentRequestNo="&sReqNo
	'Response.Write "<p>sQuery="&sQuery
	con.execute(sQuery)

	sQuery="Update Acc_T_PaymentRequestHdr set ApprovedBy="&getUserid&",ApprovedOn=getdate() where PaymentRequestNo="&sReqNo
	con.execute(sQuery)

	sQuery="insert into Acc_T_VoucherApprovalTracking (TransactionNumber,PaymentRequestNo,ApproverLevel,"&_
			"ApprovedBy,ApprovedOn) values (NULL,"&sReqNo&","&iAppLevel&","&getUserId&",getdate())"

	con.execute(sQuery)		
		
	sQuery="delete Acc_T_VouchersForApproval where PaymentRequestNo="&sReqNo&" and "&_
			"ToBeApprovedBy="&getUserId &" and ApprovalLevel="&iAppLevel
	con.execute(sQuery)
	
	if con.Errors.count <>0 then
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) &"<br>"
		next
		'Redirect to Error Handling System
	else
		con.CommitTrans
		'con.RollbackTrans
		'Response.End 
	End IF
	
Next
	%>
<SCRIPT>
	<!--
		function msgbox(strr)
		 {
				alert(strr);
				window.location.href = "PaymentRequests.asp";
	     }
	//-->
</SCRIPT>
<BODY BGCOLOR="" onLoad = "msgbox('Selected Payment Request is Approved')">
	
</BODY>
</HTML>

