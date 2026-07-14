
<%
	'Program Name				:	tempItemUpdation.asp
	'Module Name				:	Inventory (Temporary Item Creation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	October 06,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/Common/ComTempItemCreation.asp"-->

<%
Dim sItemCode

con.beginTrans

	sItemCode = Request.QueryString("ItemCode")
	
	TemporaryItemInsert(sItemCode)

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
	
	Response.Redirect "TEMPORARYITEMS.asp" 
end if

con.close
set con = nothing

%>