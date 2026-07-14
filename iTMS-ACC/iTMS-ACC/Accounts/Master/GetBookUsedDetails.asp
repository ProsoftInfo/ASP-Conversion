<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GetBookUsedDetails.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 23,2010
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
	dim objRs,objRs1,sQuery,sorgID,iBookNo,iBookid,iUseable,sCount
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs1 = Server.CreateObject("ADODB.Recordset")
	
	sorgID = Request.QueryString("OrgCode")
	iBookNo= Request.QueryString("BookNo")
	iBookid= Request.QueryString("BookCode")
	sCount = 0
	
	sQuery = "Select Count(CreatedTransNo) from ACC_T_VoucherHeader where OUDefinitionID = '"& sorgID &"' and BookCode = "& iBookid  &" and BookNumber ="& iBookNo
	Response.Write  sQuery 
	objRs.Open sQuery,con
	if not objRs.EOF then
		sCount = objRs(0)
	end if
	objRs.Close 
	
	if cstr(sCount)="0" then
	
		sQuery = "Select Count(CreatedTransNo) from ACC_T_CreatedVoucherHeader where OUDefinitionID = '"& sorgID &"' and BookCode = "& iBookid  &" and BookNumber ="& iBookNo
		Response.Write  sQuery 
		objRs.Open sQuery,con
		if not objRs.EOF then
			sCount  = objRs(0)
		end if
		objRs.Close 
	end if 'if cstr(sCount)="0" then
	Response.Clear
	if cint(sCount)>0 then
		Response.Write "Y" 
	else
		Response.Write "N"
	end if
%>
