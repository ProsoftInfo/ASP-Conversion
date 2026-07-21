<%@ Language=VBScript %>
<%	option explicit	%>
<%	Response.Expires = -10 %>
<%
	'Program Name				:	receiptNewInsert.asp
	'Module Name				:	Inventory (Receipt Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 01, 2003
	'Modified By				:	TAJUDEEN S, KUMAR K A
	'Modified On				:	August 11, 2004
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	receiptEntry.asp
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
	Dim dtCurr
	
	dtCurr = Request.QueryString("CurrDate")
	
	con.beginTrans
	
	CreateInternalReceipt dtCurr

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
