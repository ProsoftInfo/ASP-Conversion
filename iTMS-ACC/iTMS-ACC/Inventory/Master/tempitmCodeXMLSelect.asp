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
	'Program Name				:	tempitmCodeXMLSelect.asp
	'Module Name				:	Inventory (Temporary Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 10, 2003
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	tempitmCreationEntry.asp
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
	dim dcrs,Root,newElem,OutData

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT COMPANYITEMCODE FROM INV_M_ITEMMASTER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	
	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("ITEM")
			newElem.setAttribute "ITMCODE", trim(dcrs(0))
			Root.appendChild newElem
		dcrs.MoveNext
		loop
	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT GENITEMCODE FROM MS_TEMPORARYITEMMASTER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	
	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("ITEM")
			newElem.setAttribute "ITMCODE", trim(dcrs(0))
			Root.appendChild newElem
		dcrs.MoveNext
		loop
	end if
	dcrs.Close

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
