<%@ Language=VBScript %>
<%	option explicit	
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	TemporaryItemDeletion.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	October 06,2011
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
<!--#include file="../../include/sessionVerify.asp"-->
<%

Dim sQuery,iTempItemCode
Dim objRs

Set objRs = Server.CreateObject("ADODB.RecordSet")

iTempItemCode = Request.QueryString("hTempItemCode")

Con.BeginTrans

'Response.Write "<p> " & iTempItemCode
if trim(iTempItemCode)<>"" then

	
	sQuery = "Select TempItemCode from Ms_TempFinalItemDetail where TempItemCode="& iTempItemCode
	'Response.Write "<p> " & sQuery
	
	with objRs
		.ActiveConnection=con
		.CursorLocation=3
		.CursorType=3
		.Source=sQuery
		.Open
	end with
															
	set objRs.ActiveConnection = nothing
															
	if not objRs.EOF then
	
		Response.Write "Temporary Item is related to Item.You can not delete"
	
	else
		
		sQuery ="Delete from Ms_TemporaryItemMaster where TempItemCode="& iTempItemCode
		'Response.Write "<p>"&sQuery
		con.execute sQuery
		
	end if 
	objRs.Close 
						
		
	
	
	
end if ' if trim(iTempItemCode)<>"" then

	
'con.rollbacktrans
'Response.End 

Response.Clear 
con.committrans
%>
