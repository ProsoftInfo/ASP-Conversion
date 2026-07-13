<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PackingTypeDelete.asp
	'Module Name				:	INVENTORY (Master)
	'Author Name				:	UmaMaheswari S
	'Created On					:	June 04, 2011
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
<!--#include file="../../include/populate.asp"-->
<%
	Dim sQuery,nPackingCode

	nPackingCode = Request.QueryString("Code")

	con.BeginTrans
			
	sQuery = "Delete From APP_M_PackingTypeSubLevel Where PackingCode  IN ("& nPackingCode &") "
	con.Execute sQuery
		
	sQuery = "Delete From APP_M_PackingType Where PackingCode IN ("& nPackingCode &") "
	con.Execute sQuery
	
	If con.Errors.count <> "0" Then
		dim iCounter
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
	Else
		'con.RollbackTrans
		'Response.End 
		con.CommitTrans
	End IF
%>