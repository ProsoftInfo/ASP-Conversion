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
	'Program Name				:	XMLSelectAccountHead.asp
	'Module Name				:	Inventory (MRS Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	October 02, 2003
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrsCreateEntry.asp
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
	dim dcrs,sSql,OutData,Root,newElem,sIssuedFor

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	
	sIssuedFor = Request("sIssuedFor")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT ACCOUNTHEAD,CONSUMPTIONDESC FROM INV_T_CONSUMPTIONHEADRELATION WHERE ISSUEDFORCODE = '" & sIssuedFor & "' ORDER BY 2"
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing
	
	Set Root = OutData.createElement("Root")
	OutData.appendChild Root
	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("AccountHead")
			newElem.setAttribute "ACCHEAD", trim(dcrs(0))
			newElem.setAttribute "CONSUM", trim(dcrs(1))
			newElem.setAttribute "ISSFOR", sIssuedFor
			newElem.setAttribute "SRC", "D"
			Root.appendChild newElem
		dcrs.MoveNext
		loop

	end if
	dcrs.Close
	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
