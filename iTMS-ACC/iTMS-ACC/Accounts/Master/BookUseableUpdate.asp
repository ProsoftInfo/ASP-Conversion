<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BookUseableUpdate.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 16,2010
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
<%
	dim dcrs,sSql,OutData,Root,newElem
	dim objRs,objRs1,sQuery,sorgID,iBookNo,iBookid,iUseable
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs1 = Server.CreateObject("ADODB.Recordset")
	
	sorgID = Request.QueryString("OrgCode")
	iBookNo= Request.QueryString("BookNo")
	iBookid= Request.QueryString("BookCode")
	iUseable= Request.QueryString("Useable")
	
	sQuery="update Acc_R_ApplicableAccountHeads set Useable="&iUseable&" where  OUDefinitionID='"&sorgID&"'"&_
			" and BookCode="&iBookid&" and BookNumber="&iBookNo
'Response.Write sQuery	
	con.Execute (sQuery)	
	
	
	sQuery = "Select Useable from Acc_R_ApplicableAccountHeads where OUDefinitionID='"&sorgID&"'"&_
			" and BookCode="&iBookid&" and BookNumber="&iBookNo
	objRs.Open sQuery,con
	if not objRs.EOF then
		Response.Write objRs(0)
	end if
	objRs.Close
%>
