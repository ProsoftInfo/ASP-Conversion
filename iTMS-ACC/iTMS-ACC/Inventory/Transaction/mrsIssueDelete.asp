
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	mrsIssueDelete.asp
	'Module Name				:	Inventory (Issue Edit)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	APRIL 18,2014
	'Modified On				:	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	ISSUEMGMT.asp
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
<!--#include file="../../include/mrsIssueDeleteCommon.asp"-->

<%
Dim sAppCallFrom,rsIssObj,sIssQuery,sPONO,sIssEntNo,iCounter
sAppCallFrom = Request("hCallFrom")
set rsIssObj= server.CreateObject("ADODB.Recordset")
sIssEntNo = Request("ISSNO")
    con.begintrans

   MrsIssueDelete(sIssEntNo)
    if con.Errors.count <> 0 then
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
		'Redirect to Error Handling System
	else
		'Response.Write "<p>sSalInvConfirm="&sSalInvConfirm
        		
	'	con.RollbackTrans
	'	Response.End
	   Response.Clear
	   con.CommitTrans
	        Response.Redirect "ISSUEMGMT.asp?ACTN=L"
	end if 'if con.Errors.count <> 0 then
%>
