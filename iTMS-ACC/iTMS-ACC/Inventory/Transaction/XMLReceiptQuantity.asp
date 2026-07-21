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
	'Program Name				:	XMLReceiptQuantity.asp
	'Module Name				:	Inventory (Receipt Entry)
	'Author Name				:	TAJUDEEN S
	'Created On					:	August 05, 2004
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

<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->

<%
	dim dcrs,sSql,OutData,Root,newElem
	dim iItem,iClass,sOrgID,iReceiptNo
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")

	Set Root = OutData.createElement("Root")
	OutData.appendChild Root

	iItem = Request.QueryString("Item")
	iClass = Request.QueryString("Class")
	sOrgID = Request.QueryString("OrgID")
	iReceiptNo = Request.QueryString("ReceiptNo")
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT QUANTITYIN FROM RCV_T_ACTUALRCPTITEMLOT A WHERE EXISTS (SELECT RECEIPTNUMBER,ENTRYNO FROM RCV_T_ACTUALRCPTITEMDET B WHERE A.RECEIPTNUMBER = B.RECEIPTNUMBER AND A.ENTRYNO = B.ENTRYNO AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "' AND B.RECEIPTNUMBER = " & iReceiptNo & ")"
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		Root.setAttribute "QUANTITYIN",dcrs(0)
	else
		Root.setAttribute "QUANTITYIN","0"
	end if
	dcrs.Close

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
