

<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLGetPartyAddress.asp
	'Module Name				:
	'Author Name				:	
	'Created On					:   Feb
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../include/DatabaseConnection.asp"-->
<%

Dim rsTemp,oDOM,rsParty
Dim sQuery,sPartyCode
Dim ndRoot,ndParty

set rsTemp  = server.CreateObject("Adodb.Recordset")
Set rsParty = server.CreateObject("Adodb.Recordset")

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

sPartyCode= trim(Request("PartyCode"))

Set ndRoot = oDOM.createElement("Root")
oDOM.appendChild ndRoot

sQuery = "Select IsNull(AddressLine1,''),IsNull(AddressLine2,''),IsNull(City,''),IsNull(State,''),IsNull(Country,''),IsNull(PhoneNos,''),IsNull(MobileNos,'') from APP_M_PartyMaster where PartyCode = "& sPartyCode
rsParty.open sQuery,con
if not rsParty.eof then
    do while not rsParty.eof
        set ndParty = oDOM.createElement("Party")
            ndParty.setAttribute "Address1",trim(rsParty(0))
            ndParty.setAttribute "Address2",trim(rsParty(1))
            ndParty.setAttribute "City",trim(rsParty(2))
            ndParty.setAttribute "State",trim(rsParty(3))
            ndParty.setAttribute "Country",trim(rsParty(4))
            ndParty.setAttribute "PhoneNo",trim(rsParty(5))
            ndParty.setAttribute "MobileNo",trim(rsParty(6))
        ndRoot.appendChild ndParty
        rsParty.moveNext
    loop
end if
rsParty.Close
Response.ContentType="text/xml"
Response.Write oDOM.xml
%>
