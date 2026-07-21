<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MRGenerationInsert.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 29, 2005
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/MRInsertCommon.asp"-->
<%
    con.begintrans
    MRInsert

	'Response.End 
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

Response.Redirect("MRSMGMTLIST.ASP?HCHECK=M")
%>
