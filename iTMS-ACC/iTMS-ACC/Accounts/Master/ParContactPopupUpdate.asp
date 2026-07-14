<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParContactPopupUpdate.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 20, 2010
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
	Dim Root
	Dim oDOM,nodHeader,ndChild
	Dim iPartyCode
	Dim sContactNo,sConPerName,sDesignation,sPerFor,sEmailID,sQuery
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	iPartyCode = Request.QueryString("PartyCode")
	
	oDOM.async = false
	oDOM.load(Request)
	
	Con.begintrans
	
	sQuery = "Delete from APP_M_PartyContactPersons where PartyCode = "& iPartyCode
	Response.Write sQuery
	con.execute sQuery
	
	Set Root = oDOM.documentElement
	if Root.hasChildNodes() then
		for each ndChild in Root.childNodes
			sContactNo = ndChild.getAttribute("No")
			sConPerName = ndChild.getAttribute("Name") 
			sDesignation = ndChild.getAttribute("Desig")
			sPerFor = ndChild.getAttribute("PersonFor")
			sEmailID = ndChild.getAttribute("Maillid")
			
			sQuery = "INSERT INTO APP_M_PartyContactPersons(PartyCode, ContactNo, ContactPersonName, "&_
					 "Designation, ContactPersonFor, ContactMailID) VALUES("& iPartyCode &","& sContactNo &","&_
					 "'"& sConPerName &"','"& sDesignation &"','"& sPerFor &"','"& sEmailID &"')"
			Response.Write sQuery
			con.execute sQuery
		next
	end if  'if Root.hasChildNodes() then
	
'	Con.rollbacktrans
'	Response.End 
	
	Response.Clear 
	con.committrans
	
	'Response.Write Root.xml
%>