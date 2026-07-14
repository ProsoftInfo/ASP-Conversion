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
	'Program Name				:	BookGLHeadXML.asp	
	'Module Name				:	Accounts (Master)
	'Author Name				:	Senthil E
	'Created On					:	December 31, 2002
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
	dim dcrs,sSql,OutData,Root,newElem
	dim objRs,objRs1
	
	dim bFalg,sQuery,sorgID,iPartyId,iAgentType
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	
	
	sorgID = Request("orgID")
	iPartyId= Request("PartyCode")
	iAgentType= Request("AgentType")
	bFalg= Request("Flag")
	
if bFalg="A" then	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	
		sQuery="select DISTINCT PartyCode,OrgnPartyCode,PartyName from VwOrgParty where OUDefinitionID='"&sorgID&"'"&_
				" and PartyCode<>"&iPartyId&" and PartySubType=1 and PartyType='"&iAgentType &"'"&_
				" and PartyCode in (select DISTINCT AgentCode from APP_R_AgentOrgParty where OUDefinitionID='"&sorgID&"'"&_
				" and PartyCode="&iPartyId&" and AgentSubType=1) order By PartyName "
	
	'Response.write sQuery
	
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	
	if not objRs.EOF then
		do while not objRs.EOF
				Set newElem = OutData.createElement("Agent")
				newElem.setAttribute "AgentId", trim(objRs(0))
				newElem.setAttribute "AgentShortName",trim(objRs(1))
				newElem.setAttribute "Selected","Y"
				newElem.Text=trim(objRs(2))
				Root.appendChild newElem
		objRs.MoveNext
		loop
	end if
	objRs.Close
	
		sQuery="select DISTINCT PartyCode,OrgnPartyCode,PartyName from VwOrgParty where OUDefinitionID='"&sorgID&"'"&_
				" and PartyCode<>"&iPartyId&" and PartyType='"&iAgentType&"' and PartySubType=1 "&_
				" and PartyCode not in (select DISTINCT AgentCode from APP_R_AgentOrgParty where OUDefinitionID='"&sorgID&"'"&_
				" and PartyCode="&iPartyId&" and AgentSubType=1) Order By PartyName "
	
	'Response.Write sQuery
	
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	
	if not objRs.EOF then
		do while not objRs.EOF
				Set newElem = OutData.createElement("Agent")
				newElem.setAttribute "AgentId", trim(objRs(0))
				newElem.setAttribute "AgentShortName",trim(objRs(1))
				newElem.setAttribute "Selected","N"
				newElem.Text=trim(objRs(2))
				Root.appendChild newElem
		objRs.MoveNext
		loop
	end if
	objRs.Close
	
	Response.ContentType="text/xml"
	Response.Write OutData.xml
else	
end if	
%>
