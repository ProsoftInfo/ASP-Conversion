<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	StkTransferInsert.asp
	'Module Name				:	Inventory (Stock Management Stock Transfer)
	'Author Name				:	Ragavendran R
	'Created On					:	May 26, 2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
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
<!--#include file="../../include/mrsIssueInsertCommon.asp"-->
<!--#include file="../../include/getCurrentDate.asp"-->
<%
con.beginTrans

MrsIssueInsert
''Added by ragav on Aug 13,2013 for Use Existing Pack Number while moveing item from one store to another
''begin
sUseExistingPackNum = "Y"
''end
sCallFrom = "MOV"
CreateInternalReceipt(getCurrentDate())

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
'	response.end
	response.clear
	con.CommitTrans
end if
response.redirect "../Master/ITEMLISTENTRY.ASP?ACTN=MI"
%>