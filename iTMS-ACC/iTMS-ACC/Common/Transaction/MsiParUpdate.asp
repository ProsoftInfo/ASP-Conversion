<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MisParUpdate.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Aug 23, 2004
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	Code
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
	Dim Objrs,sQuery,iCtr,sParName,sParCode,sParAdd1,sParAdd2,sCity,sPin,sState,sCountry
	Dim sEMail,sItPan,sPhone,sFax,sMobile,sUrl,iPartyCode,sPanNo,sCreteBy,oDOM,sExp,ParNode
	Dim Root
	
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	oDOM.async = false
	oDOM.load(Request)
	Set Root = oDOM.documentElement
	
	sExp = "//Party"
	Set ParNode = Root.selectNodes(sExp)
	
	
	sParName=trim(ParNode.Item(0).Attributes.Item(0).nodeValue)
	sParCode=trim(ParNode.Item(0).Attributes.Item(1).nodeValue)
	sParAdd1=trim(ParNode.Item(0).Attributes.Item(2).nodeValue)
	sParAdd2=trim(ParNode.Item(0).Attributes.Item(3).nodeValue)
	sCity=trim(ParNode.Item(0).Attributes.Item(4).nodeValue)
	sPin=trim(ParNode.Item(0).Attributes.Item(5).nodeValue)
	
	sState=trim(ParNode.Item(0).Attributes.Item(6).nodeValue)
	sCountry=trim(ParNode.Item(0).Attributes.Item(7).nodeValue)
	sEmail=trim(ParNode.Item(0).Attributes.Item(8).nodeValue)
	sPanNo=trim(ParNode.Item(0).Attributes.Item(9).nodeValue)
	sPhone=trim(ParNode.Item(0).Attributes.Item(10).nodeValue)
	sFax=trim(ParNode.Item(0).Attributes.Item(11).nodeValue)
	sMobile=trim(ParNode.Item(0).Attributes.Item(12).nodeValue)
	sUrl=trim(ParNode.Item(0).Attributes.Item(13).nodeValue)
	
	sCreteBy = session("userid")
	IF Cstr(sFax) = "" Then
		sFax = "0"
	End IF
	
	Con.BeginTrans
	
	sQuery = "Select isNull(Max(MiscPartyCode),0) + 1 From App_M_MiscPartyMaster "
	Objrs.Open sQuery,Con
	IF Not Objrs.EOF Then
		iPartyCode = Objrs(0)
	End IF
	Objrs.Close
	
	sQuery = "INSERT INTO App_M_MiscPartyMaster (MiscPartyCode, OrgnPartyCode, PartyName, AddressLine1, "&_
			 "AddressLine2, City, State, Country, PhoneNos, MobileNos, FaxNos, Email, WebsiteURL, "&_
			 "Pincode, CreatedBy, CreatedOn, IncomeTaxPANNo) "&_
			 "VALUES ("&iPartyCode&", '"&sParCode&"', '"&sParName&"', '"&sParAdd1&"', '"&sParAdd2&"', '"&sCity&"', "&_
			 "'"&sState&"', '"&sCountry&"', '"&sPhone&"', '"&sMobile&"', "&sFax&", "&_
			 "'"&sEMail&"', '"&sUrl&"', '"&sPin&"', "&sCreteBy&", getDate(), '"&sPanNo&"') "
			 
	Con.Execute sQuery
	
	'Response.Write sQuery
	'con.rollbackTrans			 
	Con.CommitTrans
	
	
	
	
%>