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
	'Program Name				:	XMLItemCode.asp
	'Module Name				:	Inventory (Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 03, 2004
	'Modified By				:	TAJUDEEN S
	'Modified On				:	July 22, 2004
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
<!--#include virtual="/include/populate.asp"-->
<%
	dim dcrs,OutData,sItmType,Root,newElem,sWho,sSearch

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")

	sItmType = Request("sIType")
	sWho = Request("sWho")
	sSearch = Request("sSearch")

	Set Root = OutData.createElement("Root")
	OutData.appendChild Root

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'All Items
		if sWho = "ALL" then
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE ORDER BY 1"
		' Item Codewise
		elseif sWho = "IC" then
			sSearch = sSearch&"%"
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE COMPANYITEMCODE LIKE " & Pack(sSearch) & " ORDER BY 1"
		' Item Namewise Starts with
		elseif sWho = "IN" then
			sSearch = sSearch&"%"
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE ITEMDESCRIPTION LIKE " & Pack(sSearch) & " ORDER BY 1"
		' Item Namewise Anywhere
		elseif sWho = "IA" then
			sSearch = "%"&sSearch&"%"
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE ITEMDESCRIPTION LIKE " & Pack(sSearch) & " ORDER BY 1"
		' Classification Anywhere
		elseif sWho = "CL" then
			sSearch = "%"&sSearch&"%"
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE GROUPNAME LIKE " & Pack(sSearch) & " ORDER BY 1"
		' Item Drawing Numberwise
		elseif sWho = "DN" then
			sSearch = sSearch&"%"
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE DRAWINGNUMBER LIKE " & Pack(sSearch) & " ORDER BY 1"
		' Item Catalogue Numberwise
		elseif sWho = "CN" then
			sSearch = sSearch&"%"
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE CATALOGUENO LIKE " & Pack(sSearch) & " ORDER BY 1"
		' MGR Numberwise
		elseif sWho = "MN" then
			sSearch = sSearch&"%"
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE MGRNO LIKE " & Pack(sSearch) & " ORDER BY 1"
		' Page Numberwise
		elseif sWho = "PN" then
			sSearch = sSearch&"%"
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE PAGENO LIKE " & Pack(sSearch) & " ORDER BY 1"
		' Position Numberwise
		elseif sWho = "PO" then
			sSearch = sSearch&"%"
			.Source = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,ISNULL(DRAWINGNUMBER,'-'),ISNULL(CATALOGUENO,'-') FROM VWALLITEMCODES WHERE POSITIONNO LIKE " & Pack(sSearch) & " ORDER BY 1"
		end if
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("ITEMDETAILS")
			newElem.setAttribute "ITMCODE", trim(dcrs(0))
			newElem.setAttribute "ITMDESC", trim(dcrs(1))
			newElem.setAttribute "DRWNO", trim(dcrs(2))
			newElem.setAttribute "CATAL", trim(dcrs(3))
			Root.appendChild newElem
		dcrs.MoveNext
		loop
	end if
	dcrs.Close

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
