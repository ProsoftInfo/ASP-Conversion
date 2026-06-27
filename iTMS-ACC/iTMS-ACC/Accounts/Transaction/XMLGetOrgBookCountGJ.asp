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
	'Program Name				:	XMLGetOrgBookCountGJ.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Ragavendran R
	'Created On					:	Feb 06, 2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!-- #include File="../../include/Accpopulate.asp" -->

<%
	dim objRs,objRs1,sQuery,OutData,Root,newElem
	dim sorgID,iBookCode,sClosingCRDR,iCount

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs1 = Server.CreateObject("ADODB.Recordset")

	sorgID = Request("orgID")
	iBookCode= Request("BkCode")
	sQuery="select BookNumber,Upper(BookName),isnull(BookAccountHead,0),OtherUnitTransaction from "&_
		"vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode="&iBookCode &" Order By BookName "

	'Response.Write sQuery
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing

	Set Root = OutData.createElement("Root")
	OutData.appendChild Root

	if not objRs.EOF then
	iCount = 0
		do while not objRs.EOF
				Set newElem = OutData.createElement("Book")
				newElem.setAttribute "BookNumber", trim(objRs(0))
				newElem.setAttribute "BookName", trim(objRs(1))
				Root.appendChild newElem
				iCount = iCount + 1
		objRs.MoveNext
		loop
		Root.setAttribute "Count",iCount
	end if
	objRs.Close
Response.ContentType="text/xml"
Response.Write OutData.xml
%>

