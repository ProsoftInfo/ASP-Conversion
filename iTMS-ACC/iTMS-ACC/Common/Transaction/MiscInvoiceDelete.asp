<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MiscInvoiceDelete.asp
	'Module Name				:	Common 
	'Author Name				:	Ragavendran R
	'Created On					:	July 17,2013
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
<!--#include File="../../include/purpopulate.asp" -->
<!--#include file="../../include/Accpopulate.asp"-->
<%
Dim sQuery,iTransNo,iApp
iApp = Request("hAppCode")

iTransNo = Request.QueryString("InvNo")
con.BeginTrans

if trim(iTransNo)<>"" then
    sQuery = "Delete from Acc_T_MiscPaymentReqDetails where MiscTransNo ="& iTransNo
    Response.write "<p>"&sQuery
    con.execute sQuery
    
    sQuery = "Delete from Acc_T_MiscPymtRequestHeader where MiscTransNo ="& iTransNo
    Response.write "<p>"&sQuery
    con.execute sQuery
end if

'Con.RollBackTrans
'Response.End 
Response.clear
Con.commitTrans
Response.Redirect "MISCINVOICES.ASP?APPCODE="& iApp

%>