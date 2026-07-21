<%@ Language=VBScript %>
<%	option explicit	%>
<%	Response.Expires = -10 %>
<%
	'Program Name				:	InternalReceiptUpdate.asp
	'Module Name				:	Inventory (Receipt Updation)
	'Author Name				:	Ragavendran R
	'Created On					:	Mar 25,2014
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	receiptInternalEntry.asp
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/InternalReceiptInsert.asp"-->
<%
	Dim dtCurr,nIntRecNo
	
	dtCurr = Request.QueryString("CurrDate")
	nIntRecNo = Request("RcptNo")
	
	con.beginTrans
	
	UpdateInternalReceipt dtCurr,nIntRecNo

	if con.Errors.count <> 0 then
		dim iCounter
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
		'Redirect to Error Handling System
	else
	'	con.RollbackTrans
	'	Response.End 
		Response.Clear 
		con.CommitTrans
	end if

	con.close
	set con = nothing
	Response.Redirect "MATERIALRECEIPTS.ASP?RCPT=A"
%>
