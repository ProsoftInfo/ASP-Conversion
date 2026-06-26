<%@ Language=VBScript %>
<%
	'Program Name				:	GetUnitsForParty.asp
	'Module Name				:	Accounts
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 09,2010
	'Modified By				:	
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

Dim sQuery,sPartyCode,iCountUnit,sUnitCode
Dim rsObj
set rsObj = Server.CreateObject("ADODB.Recordset")

sPartyCode = Request.QueryString("ParCode")

sQuery = "Select distinct OUDefinitionID From APP_R_OrgParty Where PartyCode = "& sPartyCode &" and isNull(Useable,'0') = '0' "
Response.Write sQuery
rsObj.Open sQuery,con
if not rsObj.EOF then
	do while not rsObj.EOF 
		iCountUnit = iCountUnit + 1
		rsObj.MoveNext 
	loop
end if
rsObj.Close 

sQuery = "Select distinct OUDefinitionID From APP_R_OrgParty Where PartyCode = "& sPartyCode &" and isNull(Useable,'0') = '0' "
rsObj.Open sQuery,con
if not rsObj.EOF then
	do while not rsObj.EOF 
		sUnitCode = sUnitCode & "," & rsObj(0)
		rsObj.MoveNext 
	loop
end if
rsObj.Close 
if trim(sUnitCode)<>"" then
	sUnitCode = mid(sUnitCode,2)
end if
Response.Clear 
Response.Write iCountUnit &":"& sUnitCode

%>
