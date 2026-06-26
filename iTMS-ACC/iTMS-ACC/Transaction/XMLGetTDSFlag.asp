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
	'Program Name				:	XMLGetTDSFlag.asp 
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	S.MAHESWARI
	'Created On					:	January 02, 2008
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	BankVoucher.asp
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
<!-- #include File="../../include/Accpopulate.asp" -->

<%
	dim objRs,objRs1,sQuery,OutData,Root,newElem
	dim iTransNo,iEntryNo

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs1 = Server.CreateObject("ADODB.Recordset")
	
	iTransNo= Request("TransNo")	
	iEntryNo = Request("EntryNo")	
	sQuery="Select  isnull(TDSFlag,'N') from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iTransNo &" and VoucherEntryNumber = "&iEntryNo &" "
	' Response.Write sQuery
	
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
		'do while not objRs.EOF
			
		'	Set newElem = OutData.CreateElement("TDS")
		'	newElem.setAttribute "TDSFlag",objRs(0) 
		'	Root.appendchild newElem	
		'objRs.MoveNext
		'loop
		Response.ContentType="text/xml"
		'Response.Write OutData.xml
		Response.Write objRs(0) 
	end if
	objRs.Close
%>

