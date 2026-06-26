<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	RepSelectionInsert.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 04,2011
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/SessionVerify.asp"-->
<!--#include file="../../include/populate.asp"-->

<%
	Dim dcrs,ndEntry,sSql,objFS,objRoot,oDOM 'XML DOM Variable
	Dim sAgentCode,	sAgentName,sAgentSName,sAdd1,sAdd2,sCity,sPinCode,sPhone,sState,sFax,sCountry
	Dim sEmail,sURL,sMobile,ndTemp,iContactNo,sCtName,sCtDesig,sCtPer,sCtMail,iCounter,ACODE
	Dim iFillLen,sCodeLen,sAgentType,sIntOrExt,sContAdd1,sContAdd2,sContCity,sContState,sContCountry,sRepArea
	Dim iLocCode,iPartyCode,iAgentEntryID,sOrgID
				
	'DOM Objects creation
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDOM.async = false
		
	Set objFS = Server.CreateObject("Scripting.FileSystemObject")
	
	iPartyCode = Request("PartyCode")
	sOrgID = Session("organizationcode")
	
	Response.Write "<font colo=red>"
	
	If objFS.fileExists(Server.MapPath("../temp/master/RepAllocation_Party_"&Session.SessionID&".xml")) then
		oDOM.load server.MapPath("../temp/master/RepAllocation_Party_"&Session.SessionID&".xml")	
	End if
	
	Set objRoot = oDOM.DocumentElement
	
	con.beginTrans
	
	Set dcrs = server.CreateObject("Adodb.Recordset")
	
	For Each ndEntry in objRoot.childNodes
		if ndEntry.nodename = "AGENT" then
			iAgentEntryID = ndEntry.getAttribute("AgentEntryID")
			sRepArea = ndEntry.getAttribute("AreaCode")
			
			sSql = "Update APP_R_OrgParty set RepAreaCode ="& sRepArea &",RepAgentEntryID="& iAgentEntryID &" where PartyCode = "& iPartyCode &" and OUDefinitionID = "& pack(sOrgID)
	        Response.Write "<p>"&sSql
	        con.execute sSql
	    end if
	Next
	
If con.Errors.count <> 0 Then
	con.RollbackTrans
	For iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	Next
	'Redirect to Error Handling System
	
Else
   ' con.rollbacktrans
   ' Response.End 
    Response.Clear 
	con.CommitTrans
	
	If objFS.FileExists(server.MapPath("../temp/master/RepAllocation_Party_"&Session.SessionID&".xml")) Then
		objFS.deleteFile(server.MapPath("../temp/master/RepAllocation_Party_"&Session.SessionID&".xml"))
	End If
End if
	
%>