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
	'Program Name				:	LotXMLSelect.asp
	'Module Name				:	Inventory (Master - Accounting)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 21, 2003
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
	dim dcrs,OutData,sorgID,Root,iLot

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")

	sorgID = Request("orgID")
	iLot = Request("iLot")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT LOTNUMBER FROM INV_T_LOCATIONLOT WHERE ORGANISATIONCODE = '" & sorgID & "' AND LOTNUMBER = '" & iLot & "'"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	Set Root = OutData.createElement("LOT")
	OutData.appendChild Root

	if not dcrs.EOF then
		Root.setAttribute "STATUS", "Y"
	else
		Root.setAttribute "STATUS", "N"
	end if
	Response.ContentType="text/xml"
	Response.Write OutData.xml
	dcrs.Close
%>
