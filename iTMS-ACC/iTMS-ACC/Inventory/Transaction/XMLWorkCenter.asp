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
	'Program Name				:	XMLWorkCenter.asp
	'Module Name				:	Inventory (Issue Additional Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	August 04, 2005
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

<%
	dim dcrs,sSql,OutData,Root,newElem,sOrgID,sWG

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	
	sOrgID = Request("sOrgID")
	sWG = Request("WG")
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT WORKCENTERCODE,WORKCENTERNAME FROM VWWORKMACHINECENTER WHERE WORKGROUP = '" & sWG & "' AND ORGANISATIONCODE = '" & sOrgID & "' ORDER BY 2"
		.ActiveConnection = con
		.Open
	end with
'Response.Write "dcrs.Source="&dcrs.Source
	set dcrs.ActiveConnection = nothing
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("WorkCenter")
			newElem.setAttribute "WCode", trim(dcrs(0))
			newElem.setAttribute "WName", trim(dcrs(1))
			Root.appendChild newElem
		dcrs.MoveNext
		loop

		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	dcrs.Close
%>
