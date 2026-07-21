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
	'Program Name				:	XMLMachineCenter.asp
	'Module Name				:	Inventory (Issue Additional Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 06, 2004
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
	dim dcrs,sSql,OutData,Root,newElem,sOrgID,sWC,sMC

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	
	sOrgID = Request("sOrgID")
	sWC = Request("WC")
		
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT DISTINCT MACHINECENTERCODE,MACHINECENTERNAME + '--' + MakeName,MAKECODE FROM VWWORKMACHINECENTER WHERE WORKCENTERCODE = '" & sWC & "' AND ORGANISATIONCODE = '" & sOrgID & "' ORDER BY 1"
		'Added on 02Aug 2007 to display Machine Model and Serial no instead of Machine Make
		.Source = "SELECT DISTINCT MACHINECENTERCODE,MACHINECENTERNAME,MakeCode,MachineModel,MachineSerialNo FROM PRD_M_MACHINECENTER WHERE WORKCENTERCODE = '" & sWC & "' AND ORGANISATIONCODE = '" & sOrgID & "' ORDER BY 1"
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("MachineCenter")
			newElem.setAttribute "CCode", trim(dcrs(0))
			newElem.setAttribute "CName", trim(dcrs(1))
			newElem.setAttribute "MKCode", trim(dcrs(2))
			newElem.setAttribute "MCModel", trim(dcrs(3))
			newElem.setAttribute "MCSerial", trim(dcrs(4))
			Root.appendChild newElem
		dcrs.MoveNext
		loop

		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	dcrs.Close
%>
