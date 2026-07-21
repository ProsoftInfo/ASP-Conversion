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
	'Program Name				:	XMLJobWorkReceiptQuantity.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	TAJUDEEN S
	'Created On					:	September 15, 2004
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
		.Source = "SELECT DISTINCT QUANTITYIN, ISNULL(PRESSMARKNO,''), ISNULL(CROPYEAR,''), ISNULL(PRESSRUNNINGNOFROM,''), ISNULL(PRESSRUNNINGNOTO,'') FROM SUB_T_SALEARPACKRECDLOT A WHERE EXISTS (SELECT ACTUALRECEIPTNO ,ACTUALRECEIPTENTRYNO FROM SUB_T_SALEARDETAILS B WHERE A.ACTUALRECEIPTNO = B.ACTUALRECEIPTNO AND A.ACTUALRECEIPTENTRYNO = B.ACTUALRECEIPTENTRYNO AND FROMITEMCODE = " & iItem & " AND FROMCLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = '" & sOrgID & "' AND B.ACTUALRECEIPTNO = " & iReceiptNo & ")"
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		Root.setAttribute "QUANTITYIN",dcrs(0)
		Root.setAttribute "PRESSMARKNO",dcrs(1)
		Root.setAttribute "CROPYEAR",dcrs(2)
		Root.setAttribute "PRNOFROM",dcrs(3)
		Root.setAttribute "PRNOTO",dcrs(4)
	else
		Root.setAttribute "QUANTITYIN","0"
		Root.setAttribute "PRESSMARKNO",""
		Root.setAttribute "CROPYEAR",""
		Root.setAttribute "PRNOFROM",""
		Root.setAttribute "PRNOTO",""
	end if
	dcrs.Close

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
