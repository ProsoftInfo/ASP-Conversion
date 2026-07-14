<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PartySubType.asp
	'Module Name				:	Common
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Feb 03,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
Dim Salrs,oDOM

Dim ndRoot,ndPartySubType
Dim sPartyCode,sQuery,sPartytype,sOrgCode
Dim iAgentCode
set Salrs = Server.CreateObject("ADODB.Recordset")
set oDOM = Server.CreateObject("Microsoft.XMLDOM")
 sPartyCode= Request.QueryString("ParCode")
sOrgCode = Request.QueryString("OrgCode")
set ndRoot = oDOM.createElement("Root")
oDOM.appendChild ndRoot
 
	if trim(sPartyCode) <> ""  then
		sQuery = "select PartyType,PartySubType,SubTypeName from VwOrgParty where OUDefinitionID='"&sOrgCode&"' and PartyCode="&sPartyCode&" "
	else
		sQuery = "select PartyType,PartySubType,SubTypeName from VwOrgParty where OUDefinitionID='"&sOrgCode&"' Group by PartyType,PartySubType,SubTypeName"
	end if
	Salrs.Open sQuery,con
	If not Salrs.EOF then
		DO while Not Salrs.EOF
			sPartytype = Salrs(0)&"|"&Salrs(1)	
			set ndPartySubType  = oDOM.createElement("Party")
				ndPartySubType.setAttribute "SubType",sPartytype
				ndPartySubType.setAttribute "Check",""
				ndPartySubType.text=salrs(2)
				ndRoot.appendChild ndPartySubType  
			Salrs.MoveNext
			
		loop
	end if
	Salrs.Close
Response.ContentType = "text/xml"
Response.Write oDOM.xml
%>