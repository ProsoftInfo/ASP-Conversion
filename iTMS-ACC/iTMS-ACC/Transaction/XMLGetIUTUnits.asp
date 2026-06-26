<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLGetIUTUnits.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU.R
	'Created On					:	June 10, 2005
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
	dim sToOrgName,sToOrgId,OutData,Root,newElem,Objrs,sOrgID,sQuery
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	
	sOrgID = Request("UnitID")
	
	
	Set Root = OutData.createElement("Root")
	OutData.appendChild Root
	
	sQuery = "Select ToOUDefinitionID,OrgUnitShortDescription  from VwIUTDetails where FromOUDefinitionID='"&sOrgId&"' and ToOUDefinitionID in ("&_
			 "Select FromOUDefinitionID from VwIUTDetails where ToOUDefinitionID='"&sOrgId&"')"

	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	End with
	Set objRs.Activeconnection = nothing
	
	Set sToOrgId = objRs(0)
	Set sToOrgName = objRs(1)

	while not objRs.EOF
		Set newElem = OutData.createElement("IUTUnits")
		newElem.setAttribute "UnitID",sToOrgId
		newElem.setAttribute "UnitName",sToOrgName
		Root.appendChild newElem
		objRs.MoveNext
	wend
	objRs.Close
		
	Response.ContentType="text/xml"
	Response.Write OutData.xml
    
%>