<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParLocationPopupUpdate.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 20,2010
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
<%
	dim Root
	Dim oDOM,objDOM,nodHeader,ndChild,ndRoot,ndChild1
	Dim iPartyCode,sAgentFlag
	Dim sLocNo,sLocName,sAddress1,sAddress2,sCity,sState,sCountry,sECCNo
	Dim sSalesLocal,sSalesCentral,sPanNo,sStatus,sQuery
	sAgentFlag = "0"
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	' Create our DOM Document Objects
	Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	iPartyCode	= Request.QueryString("PartyCode")
	
	con.begintrans
	
		sQuery = " Delete from APP_M_AgentLocations where AgentCode = "& iPartyCode
		Response.Write sQuery
		con.execute sQuery
	
		sQuery = " Delete from APP_M_PartyLocations where PartyCode = "& iPartyCode
		Response.Write sQuery
		con.execute sQuery
	
	oDOM.async = false
	oDOM.load(Request)
	
	objDOM.load(Server.MapPath("../temp/master/Party_Master_"&Session.SessionID&".xml"))
	
	set ndRoot = objDOM.documentElement
		if ndRoot.hasChildNodes() then
			for each ndChild1 in ndRoot.childNodes
				if ndChild1.nodeName="Agent" then
					sAgentFlag = ndChild1.getAttribute("Flag")
				end if 'if ndChild1.nodeName="Agent" then
			next
		end if 'if ndRoot.hasChildNodes() then
	
		'	Response.Write ndRoot.xml
	
	Set Root = oDOM.documentElement
		' Response.Write Root.xml
	if Root.hasChildNodes() then
		for each ndChild In Root.childNodes
			sLocNo        = ndChild.getAttribute("No")
			sLocName      = ndChild.getAttribute("Name")
			sAddress1     = ndChild.getAttribute("Address1")
			sAddress2     = ndChild.getAttribute("Address2")
			sCity	      = ndChild.getAttribute("City")
			sState        = ndChild.getAttribute("State")
			sCountry      = ndChild.getAttribute("Country")
			sECCNo        = ndChild.getAttribute("ECCNo")
			sSalesLocal   = ndChild.getAttribute("SalesLocal")
			sSalesCentral = ndChild.getAttribute("SalesCentral")
			sPanNo        = ndChild.getAttribute("PANNo")
			sStatus       = ndChild.getAttribute("Status")
			
			
			sQuery = "INSERT INTO APP_M_PartyLocations(PartyCode, LocationCode,"&_
			" Location, LocationAddress1, LocationAddress2, City, State, Country,"&_
			" LocalSTNoandDT, CentralSTNoandDT, ExciseControlNo, IncomeTaxPaNNo)"&_
			" VALUES("& iPartyCode &","& sLocNo &",'"& sLocName &"','"& sAddress1 &"',"&_
			"'"& sAddress2 &"','"& sCity &"','"& sState &"','"& sCountry &"','"& sSalesLocal &"',"&_
			"'"& sSalesCentral &"','"& sECCNo &"','"& sPanNo &"')"
	
			Response.Write sQuery
			con.execute sQuery
			
			if sAgentFlag = "1" then
				sQuery = "INSERT INTO APP_M_AgentLocations(AgentCode, LocationCode, Location,"&_
					     "LocationAddress1,LocationAddress2, City, State, Country )"&_
					     "VALUES("& iPartyCode &","& sLocNo  &",'"& sLocName &"','"& sAddress1 &"',"&_
					     "'"& sAddress2 &"','"& sCity &"','"& sState &"','"& sCountry &"')"
					     
				Response.Write sQuery
				con.execute sQuery
			end if
			
		next
	end if
	
'	con.rollbacktrans
'	Response.End 
	
	Response.Clear 
	con.committrans
%>