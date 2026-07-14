<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParCreate_Edit_EntryInsert.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 17,2010
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
dim sName,sShortName,sAddress1,sAddress2,sPincode,sCity,sState,sCountry,sPhone,sFax,sAction
dim sEmail,sWebsite,sECCNo,sSalesLocal,sSalesCentral,sGroupFlag,sGroup
DIM Temp,iCounter,arrUnit,arrUnitName,sPanNo,sMobile,iPartyCode,sSelUnit,iCreatedBy
Dim oDOM,Root,newElem,newElem1,sCallTy,Temparr,sDelType,sOwnUnit,ndChild
dim sQuery,iRecCount,objRs,sHisno,sHisRes,sTinNumber,sInActive,sAgent,rsTemp
Dim iAgentEntryID

Dim ndUnit,ndUN,ndPartyType,sOpeningAmt,sOpCRDR,sOpenMonthYear,sCloseMonthYear
Dim sUnitCode,sPartyType,sPartySubType,sParentPartyCode,sQuery1

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set rsTemp = Server.CreateObject("ADODB.RecordSet")



iCreatedBy = getUserid()

iPartyCode = Request.QueryString("PartyCode")
'Response.Write "iPartyCode = "& iPartyCode
sAction = Request.QueryString("Action")

Response.Write "sAction = "& sAction

if trim(iPartyCode)<>"" then

	sQuery ="Delete from Acc_T_PartyOpeningAmt where PartyCode="& iPartyCode
	Response.Write "<p>"&sQuery
	con.execute sQuery
		                
	sQuery ="Delete from APP_R_OrgParty where PartyCode="& iPartyCode
	Response.Write "<p>"&sQuery
	con.execute sQuery
	
end if ' if trim(iPartyCode)<>"" then

con.begintrans

oDOM.load(server.MapPath("../temp/master/Party_Master_"&Session.SessionID&".xml"))

	Set Root = oDOM.documentElement
	
	
	if Root.hasChildNodes() then
		for each ndChild in Root.childNodes
			if ndChild.nodeName="ParName" then
				sName = ndChild.text
			elseif ndChild.nodeName="ShortName" then
				sShortName = ndChild.text
			elseif ndChild.nodeName="Address1" then
				sAddress1 = ndChild.text
			elseif ndChild.nodeName="Address2" then
				sAddress2 = ndChild.text
			elseif ndChild.nodeName="PinCode" then
				sPincode = ndChild.text
			elseif ndChild.nodeName="City" then
				sCity = ndChild.text
			elseif ndChild.nodeName="State" then
				sState = ndChild.text
			elseif ndChild.nodeName="Country" then
				sCountry= ndChild.text
			elseif ndChild.nodeName="Phone" then
				sPhone = ndChild.text
			elseif ndChild.nodeName="Mobile" then
				sMobile = ndChild.text
			elseif ndChild.nodeName="Fax" then
				sFax = ndChild.text
			elseif ndChild.nodeName="Email" then
				sEmail = ndChild.text
			elseif ndChild.nodeName="Website" then
				sWebsite = ndChild.text
			elseif ndChild.nodeName="ECCNo" then
				sECCNo = ndChild.text
			elseif ndChild.nodeName="PANNo" then
				sPanNo = ndChild.text
			elseif ndChild.nodeName="CreatedBy" then
				iCreatedBy = ndChild.text
			elseif ndChild.nodeName="OwnUnit" then
				sOwnUnit = ndChild.text
			elseif ndChild.nodeName="Sales" then
				sSalesLocal = ndChild.getAttribute("Local")
				sSalesCentral = ndChild.getAttribute("Central")
			elseif ndChild.nodeName="TINNumber" then
				sTinNumber= ndChild.text
			elseif ndChild.nodeName="Group" then
				sGroupFlag = ndChild.getAttribute("Flag")
				sGroup = ndChild.getAttribute("Type")
				sParentPartyCode = ndChild.getAttribute("ParentCompany")
			elseif ndChild.nodeName="Active" then
				sInActive = ndChild.getAttribute("Flag")
			elseif ndChild.nodeName="Agent" then
				sAgent=ndChild.getAttribute("Flag")
			end if
		next
	end if
	
	sName = mid(sName,1,100)
if trim(sAction)="Edit" then

	if trim(sOwnUnit)="" then
		sOwnUnit = "NULL"
	else
		sOwnUnit = pack(sOwnUnit)
	end if
	
	sQuery= "Update APP_M_PartyMaster set OrgnPartyCode='"&sShortName&"',PartyName='"&sName&"',"&_
			"AddressLine1='"&sAddress1&"',AddressLine2='"&sAddress2&"',City='"&sCity&"',"&_
			"State='"&sState&"',Country='"&sCountry&"',PhoneNos='"&sPhone&"',MobileNos='"&sMobile&"',"&_
			"FaxNos='"&sFax&"',Email='"&sEmail&"',WebsiteURL='"&sWebsite&"',Pincode='"&sPincode&"',"&_
			"ExciseControlCode='"&sECCNo&"',LocalSTNoandDT='"&sSalesLocal&"',CentralSTNoandDT='"&sSalesCentral&"',"&_
			"IncomeTaxPANNo='"&sPanNo&"',TINNumber='"&sTinNumber&"',ParentPartycode="& sParentPartyCode &","&_
			"OwnUnitCode="& sOwnUnit &",Useable="& sInActive &",PartyGroupCoyType='"& sGroup &"' where PartyCode = "& iPartyCode
		
	Response.Write "<p>"&sQuery
	Con.Execute sQuery
	
	if trim(sOwnUnit)<>"" then
		sQuery = "UPDATE DCS_ORGANIZATIONUNITDEFINITIONS SET PARTYCODE = "& iPartyCode &", PARTYMAPPEDBY = "& iCreatedBy &","&_
				 "PARTYMAPPEDON = getdate() WHERE OUDEFINITIONID= "& sOwnUnit
		Response.Write "<p>"&sQuery
		Con.Execute sQuery
	end if	'if trim(sOwnUnit)<>"" then	 
	
	
	if trim(sAgent)="1" then
	
	    sQuery = "Select isNull(max(AgentEntryID),0)+1 from APP_M_AgentMaster"
	    rsTemp.Open sQuery,con
	    if not rsTemp.EOF then
	        iAgentEntryID = rsTemp(0)
	    end if
	    rsTemp.Close 
		
	    sQuery = "Select * from App_M_AgentMaster where AgentCode = "& iPartyCode
	    Response.Write "<p>"&sQuery
	    rsTemp.open sQuery,con
	    if rsTemp.eof then
	        sQuery = "INSERT INTO APP_M_AgentMaster(AgentCode, AgentShortName, AgentName,"&_
					 "AgentAddress1,AgentAddress2,AgentAddress3, AgentAddress4, AgentPhone,"&_
					 "AgentFax, AgentMailID,AgentEntryID) VALUES("& iPartyCode &",'"& sShortName &"','"& sName &"',"&_
					 "'"& sAddress1 &"','"& sAddress2 &"','"& sCity &"'+'-'+'"& sPincode &"','"& sState&"'+'  '+'"&sCountry&"',"&_
					 "'"& sPhone &"','"& sFax &"','"&sEmail&"',"& iAgentEntryID  &")"
    
    		sQuery1 = "INSERT INTO APP_M_AgentLocations(AgentCode, LocationCode, Location, City,AgentEntryID )"&_
				     "VALUES("& iPartyCode &",1,'"& sCity &"','"& sCity &"',"& iAgentEntryID  &")"
					     
	    else
	        sQuery = "UPDATE APP_M_AgentMaster set AgentShortName='"& sShortName &"',AgentName='"& sName &"',"&_
        		 "AgentAddress1='"& sAddress1 &"',AgentAddress2='"& sAddress2 &"',AgentAddress3='"& sCity &"'+'-'+'"& sPincode &"',"&_
				 "AgentAddress4='"& sState&"'+'  '+'"&sCountry&"', AgentPhone='"& sPhone &"',"&_
				 "AgentFax='"& sFax &"', AgentMailID='"&sEmail&"' where AgentCode = "& iPartyCode
	    end if
	    rsTemp.Close 
		Response.Write "<p>"&sQuery
		Con.Execute sQuery 
		if Trim(sQuery1)<>"" then
		    Response.Write "<p>"&sQuery1
		    con.execute sQuery1
		end if 'if Trim(sQuery1)<>"" then
	else
		sQuery = "Delete from APP_M_AgentMaster where AgentCode = "& iPartyCode
		Response.Write "<p>"&sQuery
		con.execute sQuery
		
		sQuery = "Delete from APP_M_AgentLocations where AgentCode = "& iPartyCode
		Response.Write "<p>"&sQuery
		con.execute sQuery
	end if
	
		if Root.hasChildNodes() then
	    for each ndUnit in Root.childNodes
	        if ndUnit.nodeName="Units" then
	            for each ndUN in ndUnit.childNodes
	                sUnitCode = ndUN.getAttribute("Code")
	                
	                for each ndPartyType in ndUN.childNodes
	                     sPartyType = ndPartyType.getAttribute("Type")
	                     sPartySubType = ndPartyType.getAttribute("SubType")
	                     sOpeningAmt = ndPartyType.getAttribute("OpBalance")
	                     sOpCRDR = ndPartyType.getAttribute("OpCRDR")
	                     sOpenMonthYear = ndPartyType.getAttribute("OpeningMonthYear")
	                     sCloseMonthYear = ndPartyType.getAttribute("ClosingMonthYear")
	                     

'				Response.Write "sOpCRDR="& sOpCRDR        
	                     sQuery = "INSERT INTO APP_R_OrgParty(OUDefinitionID, PartyType, PartySubType,"&_
	                              "PartyCode, OpeningBalance, OpeningCDIndication, PrefTransporterCode,"&_
	                              "PrefDespatchMode, PrefCurrencyCode, PrefPaymentMode, PrefBasisOfPricing,"&_
	                              "PrefPaymentTerms) VALUES('"& sUnitCode &"','"& sPartyType &"',"& sPartySubType &","&_
	                              ""& iPartyCode &","& sOpeningAmt &",'"& sOpCRDR &"',0,0,0,0,0,0)"
	                              
	                     Response.Write "<p>"&sQuery
	                     con.execute sQuery
	                     
	                     sQuery = "INSERT INTO Acc_T_PartyOpeningAmt(OUDefinitionID, PartyType, PartySubType,"&_
	                              "PartyCode, OpeningMonthYear,OpeningAmount,OpeningCDIndication,ClosingMonthYear)"&_
	                              " VALUES('"& sUnitCode &"','"& sPartyType &"',"& sPartySubType &","& iPartyCode &","&_
	                              "'"& sOpenMonthYear &"',"& sOpeningAmt &",'"& sOpCRDR &"','"& sCloseMonthYear &"')"
	                              
	                     Response.Write "<p>"&sQuery
	                     con.execute sQuery
	                     
	                next
	            next
	        end if 'if ndUnit.nodeName="Units" then
	    next	    
	end if 'if Root.hasChildNodes() then
	
else

	sQuery = "Select isNull(Max(PartyCode),0)+1 from APP_M_PartyMaster"
	objRs.Open sQuery,con 
	if not objRs.EOF then
		iPartyCode = objRs(0)
	end if 
	objRs.Close 
	
	if trim(sOwnUnit)="" then
		sOwnUnit = "NULL"
	else
		sOwnUnit = pack(sOwnUnit)
	end if


	sQuery = "Insert into APP_M_PartyMaster (PartyCode,OrgnPartyCode ,PartyName,AddressLine1,AddressLine2,"&_
	"City,State,Country,PhoneNos,MobileNos,FaxNos,Email,WebsiteURL,Pincode,ExciseControlCode,"&_
	"LocalSTNoandDT,CentralSTNoandDT,IncomeTaxPANNo,TINNumber,CreatedBy, CreatedOn, ApprovedBy, ApprovedOn,ParentPartycode,OwnUnitCode,PartyGroupCoyType)"&_
	" values ("&iPartyCode&",'"&sShortName&"','"&sName&"','"&sAddress1&"','"&sAddress2&"', '"&sCity&"','"&sState&"', '"&sCountry&"','"&sPhone&"','"&_
	""&sMobile&"','"&sFax&"','"&sEmail&"','"&sWebsite&"', '"&sPincode&"',"&_
	"'"&sECCNo&"','"&sSalesLocal&"','"&sSalesCentral&"','"&sPanNo&"','"&sTinNumber&"',"& iCreatedBy &",getdate(),"& iCreatedBy &",getdate(),"& sParentPartyCode &","& sOwnUnit &",'"& sGroup &"')"
	
	Response.Write "<p>"&sQuery
	Con.Execute sQuery
	
	
	if trim(sOwnUnit)<>"" then
		sQuery = "UPDATE DCS_ORGANIZATIONUNITDEFINITIONS SET PARTYCODE = "& iPartyCode &", PARTYMAPPEDBY = "& iCreatedBy &","&_
				 "PARTYMAPPEDON = getdate() WHERE OUDEFINITIONID= "& sOwnUnit
		Response.Write "<p>"&sQuery
		Con.Execute sQuery
	end if	'if trim(sOwnUnit)<>"" then	 
	
	if trim(sAgent)="1" then
	    sQuery = "Select isNull(max(AgentEntryID),0)+1 from APP_M_AgentMaster"
	    rsTemp.Open sQuery,con
	    if not rsTemp.EOF then
	        iAgentEntryID = rsTemp(0)
	    end if
	    rsTemp.Close 
	
			sQuery = "INSERT INTO APP_M_AgentMaster(AgentCode, AgentShortName, AgentName,"&_
					 "AgentAddress1,AgentAddress2,AgentAddress3, AgentAddress4, AgentPhone,"&_
					 "AgentFax, AgentMailID,AgentEntryID) VALUES("& iPartyCode &",'"& sShortName &"','"& sName &"',"&_
					 "'"& sAddress1 &"','"& sAddress2 &"','"& sCity &"-"& sPincode &"','"& sState&"-"&sCountry&"',"&_
					 "'"& sPhone &"','"& sFax &"','"&sEmail&"',"& iAgentEntryID &")"
			Response.Write "<p>"&sQuery
			con.execute sQuery
	end if
	
	
	sQuery = "INSERT INTO APP_M_PartyLocations(PartyCode, LocationCode,Location, City)"&_
			" VALUES("& iPartyCode &",1,'"& sCity &"','"& sCity &"')"
	
		Response.Write "<p>"&sQuery
		con.execute sQuery
			
		if trim(sAgent) = "1" then
			sQuery = "INSERT INTO APP_M_AgentLocations(AgentCode, LocationCode, Location, City,AgentEntryID )"&_
				     "VALUES("& iPartyCode &",1,'"& sCity &"','"& sCity &"',"& iAgentEntryID &")"
					     
			Response.Write "<p>"&sQuery
			con.execute sQuery
		end if
	
	if Root.hasChildNodes() then
	    for each ndUnit in Root.childNodes
	        if ndUnit.nodeName="Units" then
	            for each ndUN in ndUnit.childNodes
	                sUnitCode = ndUN.getAttribute("Code")
	                for each ndPartyType in ndUN.childNodes
	                     sPartyType = ndPartyType.getAttribute("Type")
	                     sPartySubType = ndPartyType.getAttribute("SubType")
	                     sOpeningAmt = ndPartyType.getAttribute("OpBalance")
	                     sOpCRDR = ndPartyType.getAttribute("OpCRDR")
	                     sOpenMonthYear = ndPartyType.getAttribute("OpeningMonthYear")
	                     sCloseMonthYear = ndPartyType.getAttribute("ClosingMonthYear")
	                     

'				Response.Write "sOpCRDR="& sOpCRDR        
	                     sQuery = "INSERT INTO APP_R_OrgParty(OUDefinitionID, PartyType, PartySubType,"&_
	                              "PartyCode, OpeningBalance, OpeningCDIndication, PrefTransporterCode,"&_
	                              "PrefDespatchMode, PrefCurrencyCode, PrefPaymentMode, PrefBasisOfPricing,"&_
	                              "PrefPaymentTerms) VALUES('"& sUnitCode &"','"& sPartyType &"',"& sPartySubType &","&_
	                              ""& iPartyCode &","& sOpeningAmt &",'"& sOpCRDR &"',0,0,0,0,0,0)"
	                              
	                     Response.Write "<p>"&sQuery
	                     con.execute sQuery
	                     
	                     sQuery = "INSERT INTO Acc_T_PartyOpeningAmt(OUDefinitionID, PartyType, PartySubType,"&_
	                              "PartyCode, OpeningMonthYear,OpeningAmount,OpeningCDIndication,ClosingMonthYear)"&_
	                              " VALUES('"& sUnitCode &"','"& sPartyType &"',"& sPartySubType &","& iPartyCode &","&_
	                              "'"& sOpenMonthYear &"',"& sOpeningAmt &",'"& sOpCRDR &"','"& sCloseMonthYear &"')"
	                              
	                     Response.Write "<p>"&sQuery
	                     con.execute sQuery
	                     
	                next
	            next
	        end if 'if ndUnit.nodeName="Units" then
	    next	    
	end if 'if Root.hasChildNodes() then

end if

'	con.rollbacktrans
'	Response.End 

	Response.Clear 
	con.committrans

	Response.Redirect "ParCreate_Edit_Entry.asp?PartyCode="&iPartyCode

%>
