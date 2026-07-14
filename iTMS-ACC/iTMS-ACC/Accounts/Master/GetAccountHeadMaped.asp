<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GetAccountHeadMaped.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 31,2012
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
	dim objRs,objRs1,sQuery,sorgID,iBookNo,iBookid,iUseable,sCount,iAccHead
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs1 = Server.CreateObject("ADODB.Recordset")
	
	iAccHead = Request.QueryString("AccHead")
	sorgID = Request.QueryString("OrgCode")
	sCount = 0
	
	sQuery = "Select Count(*) from Acc_R_ApplicableAccountHeads where BookAccountHead = "& iAccHead &" and OUDefinitionID = '"& sorgID &"' "
	Response.Write  sQuery 
	objRs.Open sQuery,con
	if not objRs.EOF then
		sCount = objRs(0)
	end if
	objRs.Close 
	Response.Clear
	Response.Write sCount
%>
