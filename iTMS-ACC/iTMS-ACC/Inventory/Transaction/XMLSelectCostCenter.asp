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
	'Program Name				:	XMLSelectCostCenter.asp
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
	dim dcrs,sSql,OutData,Root,newElem,sOrgID

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	
	sOrgID = Request("sOrgID")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT COSTCENTERHEAD,CCHEADCODE FROM VWORGCOSTCENTER WHERE OUDEFINITIONID = '" & sOrgID & "' ORDER BY COSTCENTERHEAD"
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("CostCenter")
			newElem.setAttribute "CCode", trim(dcrs(0))
			newElem.setAttribute "CName", trim(dcrs(1))
			Root.appendChild newElem
		dcrs.MoveNext
		loop

		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	dcrs.Close
%>
