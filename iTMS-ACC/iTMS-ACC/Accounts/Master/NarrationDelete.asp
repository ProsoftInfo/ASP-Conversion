<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	NarrationDelete.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	March 28, 2011
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
	Dim sQuery,objRs,nNarrationNo,nBookCode
	
	nNarrationNo = Request.QueryString("NarrationNo")
	nBookCode = Trim(Request.QueryString("BookCode"))
	
	Set objRs = Server.CreateObject("ADODB.RecordSet")


	Con.Begintrans

	sQuery="delete  from  Acc_R_BookFreqDesc where NarrationNumber IN ("& nNarrationNo &" ) "
	con.Execute(sQuery)

	sQuery="delete from  Acc_M_FrequentDescriptions where NarrationNumber IN ("& nNarrationNo &" ) "
	con.Execute(sQuery)

	'con.RollbackTrans
	'Response.End 
	
	Con.CommitTrans
	con.close
	Set con = Nothing
	
	Response.Redirect ("BOOKNARRATIONS.ASP?BOOKCODE="&nBookCode)
%>
