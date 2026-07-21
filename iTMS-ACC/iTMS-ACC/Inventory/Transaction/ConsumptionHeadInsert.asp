<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ConsumptionHeadInsert.asp
	'Module Name				:	Inventory (Issue)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 08, 2007
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<%
	dim newxml
	dim dcrs,sSql,RootNode,HeaderNode
	dim sExp,iCtr,iAHead,sDesc,sIssFor
	
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	' Create our DOM Document Objects
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")

	newxml.async = false
	newxml.load(Request)

	Set RootNode = newxml.documentElement

	con.beginTrans

	sExp ="//AccountHead [ @SRC = 'N']"
	Set HeaderNode = RootNode.Selectnodes(sExp)
	if HeaderNode.Length > 0 then
		For iCtr = 0 to HeaderNode.Length - 1
			iAHead = trim(HeaderNode.Item(iCtr).Attributes.getNamedItem("ACCHEAD").Value)
			sDesc = trim(HeaderNode.Item(iCtr).Attributes.getNamedItem("CONSUM").Value)
			sIssFor = trim(HeaderNode.Item(iCtr).Attributes.getNamedItem("ISSFOR").Value)

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT DISTINCT ACCOUNTHEAD FROM INV_T_CONSUMPTIONHEADRELATION WHERE ISSUEDFORCODE = '" & sIssFor & "' AND LOWER(CONSUMPTIONDESC) = '" & lcase(sDesc) & "' AND ACCOUNTHEAD = " & iAHead & ""
				.ActiveConnection = con
				.Open
			end with

			set dcrs.ActiveConnection = nothing
			if dcrs.EOF then
				sSql = "INSERT INTO INV_T_CONSUMPTIONHEADRELATION (ISSUEDFORCODE,CONSUMPTIONDESC,ACCOUNTHEAD) VALUES " &_
					"('" & sIssFor & "','" & sDesc & "'," & iAHead & ")"
				Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs.Close

		next
	end if

	if con.Errors.count <> 0 then
		dim iCounter
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
		'Redirect to Error Handling System
	else
	'	con.RollbackTrans
	'	Response.End
		con.CommitTrans
	end if

	con.close
	set con = nothing

%>
