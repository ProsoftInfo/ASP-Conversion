<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PaymtReqDeletion.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	05th April 2011
	'Modified By				:   
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
dim sQuery,sReqNo,sReqStatus,sActionFlag,iAppLevel

sReqNo=Request("RequestNo")
con.BeginTrans

sQuery="Delete From Acc_T_PaymentRequestDet where PaymentRequestNo IN ("&sReqNo&") "
con.execute(sQuery)

sQuery="Delete From Acc_T_PaymentRequestHdr where PaymentRequestNo IN ("&sReqNo &")"
con.execute(sQuery)

sQuery="Delete From Acc_T_VoucherApprovalTracking  where PaymentRequestNo IN ("&sReqNo &") "
con.execute(sQuery)		
	
sQuery="delete From Acc_T_VouchersForApproval where PaymentRequestNo IN ("&sReqNo&") "
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
	<BODY BGCOLOR="" onLoad = "msgbox('Selected Payment is deleted Successfully')">
	<%
end if
%>
</BODY>
</HTML>

