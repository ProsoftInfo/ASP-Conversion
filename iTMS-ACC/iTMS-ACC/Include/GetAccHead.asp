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
	'Program Name				:	GetAccHead.asp
	'Module Name				:	
	'Author Name				:	UmaMaheswari S
	'Created On					:	May 05, 2011
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

<!--#include file="DatabaseConnection.asp"-->
<!--#include file="Accpopulate.asp"-->
<%
	Dim iAccHead,oDom,sTemp,sUnitID,sPartyType,sPartySubType
	
	sUnitID = Session("organizationcode")
	
	sTemp = trim(Request.QueryString("Data"))
	sPartyType = Split(sTemp,":")(0)
	sPartySubType = Split(sTemp,":")(1)
	
	iAccHead = GetCtrlAccForParty(sPartyType,sPartySubType,sUnitID)
	
	Response.ContentType="text/xml"
	Response.Write iAccHead
		
%>