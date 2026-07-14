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
	'Program Name				:	XMLAssetItem.asp
	'Module Name				:	Inventory (Asset Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 01, 2004
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
	dim dcrs,sSql,OutData,Root,newElem,sOrgID,iClass

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	
	sOrgID = Request("sOrgID")
	iClass = Request("sClass")
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT DISTINCT ITEMCODE,ITEMDESCRIPTION,STORESUOM FROM VWALLITEMS WHERE CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "' ORDER BY 2"
		.Source = "SELECT DISTINCT ITEMCODE,ITEMDESCRIPTION,STORESUOM FROM VW_INV_ITEMS WHERE CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "' ORDER BY 2"
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("ITEM")
			newElem.setAttribute "CLACODE", iClass
			newElem.setAttribute "ITMCODE", trim(dcrs(0))
			newElem.setAttribute "INAME", trim(dcrs(1))
			newElem.setAttribute "SUOM", trim(dcrs(2))
			Root.appendChild newElem
		dcrs.MoveNext
		loop

		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	dcrs.Close
%>
