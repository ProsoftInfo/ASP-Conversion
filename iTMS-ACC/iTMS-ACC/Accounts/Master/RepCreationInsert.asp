<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	RepCreationInsert.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 03,2011
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/SessionVerify.asp"-->
<!--#include virtual="/include/populate.asp"-->

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
	iAgentEntryId = Request("AgentEntryID")
    sOrgID = Session("organizationcode")
	
	
	Response.Write "<font colo=red>"
	
	If objFS.fileExists(Server.MapPath("../temp/master/Rep_Party_"&Session.SessionID&".xml")) then
		oDOM.load server.MapPath("../temp/master/Rep_Party_"&Session.SessionID&".xml")	
	End if
	
	Set objRoot = oDOM.DocumentElement
	
	con.beginTrans
	
	Set dcrs = server.CreateObject("Adodb.Recordset")
	if Trim(iAgentEntryID)="0" then iAgentEntryID = ""
	
	if Trim(iAgentEntryID)="" then
	    'To Autogenerate Agent code
	    sSql = "Select isnull(str(max(AgentEntryID)),0)+1 from APP_M_AgentMaster"
	    With dcrs 
		    .CursorLocation = 3
		    .CursorType = 3
		    .Source =  sSql
		    .ActiveConnection = con
		    .Open
	    End With
	    Set dcrs.ActiveConnection = nothing
    		
	    if not dcrs.EOF then
		    ACODE = trim(dcrs(0))
	    end if
	    dcrs.Close 
	else
	    ACODE = iAgentEntryID
	end if
	
	if Trim(iAgentEntryID)<>"" then
		    sSql ="Delete from APP_M_AgentLocations where AgentEntryID = "& ACODE
		    Response.Write "<p>"& sSql
		    con.execute sSql
		    
		    sSql = "Delete from APP_M_Agentmaster where AgentEntryID = "& ACODE
		    Response.Write "<p>"& sSql
		    con.execute sSql
    end if

	For Each ndEntry in objRoot.childNodes
		if ndEntry.nodename = "AGENTHEADER" then
		
			sAgentName = ndEntry.getAttribute("AgentName")
			sAgentSName	= ndEntry.getAttribute("AgentSName")
			sAdd1	= ndEntry.getAttribute("Add1")
			sAdd2	= ndEntry.getAttribute("Add2")
			sCity	= ndEntry.getAttribute("City")
			sPinCode =	ndEntry.getAttribute("PinCode")
			sPhone	= ndEntry.getAttribute("Phone")
			sState	= ndEntry.getAttribute("State")
			sFax	= ndEntry.getAttribute("Fax")
			sMobile =	ndEntry.getAttribute("Mobile")
			sCountry =	ndEntry.getAttribute("Country")
			sEmail	= ndEntry.getAttribute("Email")
			sURL	= ndEntry.getAttribute("URL")
			sAgentType = ndEntry.getAttribute("AgentType")
			sIntOrExt = ndEntry.getAttribute("IntOrExt")
			
		    sSql = "Insert into APP_M_Agentmaster (AgentEntryID,AgentName,AgentShortName,"&_
		           " AgentAddress1,AgentAddress2,AgentAddress3,AgentAddress4,AgentPhone,"&_
		           " AgentFax,AgentMailID,ExternalOrInternal,AgentType)"&_
		           " values("& ACODE &","&Pack(sAgentName)&","& Pack(sAgentSName) &","&_
		           " "& Pack(sAdd1) &","& Pack(sAdd2) &","& Pack(sCity &"-"& sPinCode) &",NULL,"&_
		           " "& Pack(sMobile) &","& Pack(sFax) &","& Pack(sEmail) &","& Pack(sIntOrExt)&","&Pack(sAgentType)&")"
	        Response.Write "<p>"& sSql 
	        con.execute sSql
	    end if
		
		if ndEntry.nodename = "CONTACTDETAIL" then
			iContactNo = 0
			For each ndTemp in ndEntry.childNodes
				iContactNo = iContactNo + 1
				sCtName  = ndTemp.getAttribute("Name")
				sCtDesig = ndTemp.getAttribute("Desig")
				sCtPer   = ndTemp.getAttribute("ContactFor")
				sCtMail  = ndTemp.getAttribute("Email")
				sContAdd1 = ndTemp.getAttribute("Address1")
	            sContAdd2 = ndTemp.getAttribute("Address2")
	            sContCity = ndTemp.getAttribute("City")
	            sContState = ndTemp.getAttribute("State")
	            sContState = ndTemp.getAttribute("Country")
	            sRepArea = ndTemp.getAttribute("RepArea")
	            
	            sSql ="Select isNull(Max(LocationCode),0)+1 from APP_M_AgentLocations where AgentEntryID = "& ACODE
	            dcrs.Open sSql,con
	            if not dcrs.EOF then
	                iLocCode = dcrs(0)
	            end if
	            dcrs.Close 
	            
	            sSql = "Insert into APP_M_AgentLocations (AgentEntryID,LocationCode,Location,"&_
	                " LocationAddress1,LocationAddress2,City,State,Country,ContactPersonName,"&_
	                " Designation,ContactPersonFor,ContactMailID,RepresentingAreaCode)"&_
	                " values("& ACODE &","& iLocCode &","& Pack(sCity) &","& pack(sContAdd1) &","&_
	                " "& Pack(sContAdd2) &","& Pack(sCity) &","& Pack(sState) &","& Pack(sCountry) &","&_
	                " "& Pack(sCtName) &","& Pack(sCtDesig) &","& Pack(sCtPer) &","& Pack(sCtMail)&","& sRepArea &")"
	           Response.Write "<p>"& sSql
	           con.execute sSql
	            
			Next
		End if
	Next
	
'	sSql = "Update APP_R_OrgParty set RepAreaCode ="& sRepArea &",RepAgentEntryID="& ACODE &" where PartyCode = "& iPartyCode &" and OUDefinitionID = "& pack(sOrgID)
'	Response.Write "<p>"&sSql
'	con.execute sSql
	
	
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
	
	If objFS.FileExists(server.MapPath("../temp/master/Rep_Party_"&Session.SessionID&".xml")) Then
		objFS.deleteFile(server.MapPath("../temp/master/Rep_Party_"&Session.SessionID&".xml"))
	End If
	Response.Redirect ("RepCreationEntry.asp?PartyCode="&iPartyCode)
End if
	
%>