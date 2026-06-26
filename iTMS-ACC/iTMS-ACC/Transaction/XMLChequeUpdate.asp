<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLChequeUpdate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	September 17,2003
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
<%
dim iTransNo,sChequeNo,sQuery

iTransNo=Request("TransNo")
sChequeNo=Request("ChequeNo")
		

sQuery="update Acc_T_VoucherHeader set ChequePrinted=1,BankInstrumentNo='"&sChequeNo&"'"&_
	" where CreatedTransNo="&iTransNo
	
con.execute(sQuery)		
sQuery="update Acc_T_CreatedVoucherHeader set BankInstrumentNo='"&sChequeNo&"'"&_
	" where CreatedTransNo="&iTransNo
		
con.execute(sQuery)
		
		
%>