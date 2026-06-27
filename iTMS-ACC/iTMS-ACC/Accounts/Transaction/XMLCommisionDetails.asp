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
	'Program Name				:	XMLGetOrgBook.asp	
	'Module Name				:	ACCOUNTS (Transcation) 
	'Author Name				:	SENTHIL E
	'Created On					:	February 27, 2003
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	VouCNBookSelection.asp
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

<!-- #include File="../../include/DatabaseConnection.asp" -->

<%
	dim OutData,Root,newElem
	dim sAgentCode,sQuery,objRs,sTemp,sOrgId,sSelInvno

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	
	sAgentCode = Request("AgentCode")
	sOrgId = Request("Orgid")
	sTemp=Split(sAgentCode,"?")
	sSelInvno = Request("sSelInv")
	
	IF CStr(sSelInvno) = "" Then
		sQuery="select AccTransactionNo,VoucherNumber,convert(char,VoucherDate,103),Str(AgentCommission,11,2) from "&_
		" VwAgentCommisionDetails where OUDefinitionID='"&sOrgId&"' and AgentCode="&Trim(sTemp(3)) &" and CommissionToPay=1"
	Else
		sQuery = "select Distinct AccTransactionNo,VoucherNumber,convert(char,VoucherDate,103), "&_
				 "Str(AgentCommission,11,2) from VwAgentCommisionDetails  "&_
				 "where AccTransactionNo in ("&sSelInvno&") and CommissionToPay = 0 and AgentCode="&Trim(sTemp(3))&" "
	End IF
	
	
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	
	if not objRs.EOF then
		do while not objRs.EOF
				Set newElem = OutData.createElement("AGCommision")
				newElem.setAttribute "TransNo", trim(objRs(0))
				newElem.setAttribute "Amount", trim(objRs(3))
				newElem.setAttribute "CommDetails", trim(objRs(1))&"--"&trim(objRs(2))&"--"&trim(objRs(3))
				Root.appendChild newElem
		objRs.MoveNext
		loop
		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	objRs.Close
%>
