<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLSaveParty.asp
	'Module Name				:	Accounts(master)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 18,2010
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
<%
	Dim oDOM,sMod,sName,sDesgDir

	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	sMod=Request("Mod")
	sName=Request("Name")
	
	oDOM.async = false
	oDOM.load(Request)
	
	oDOM.Save server.MapPath("../temp/master/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")	
	
	
	set oDOM=nothing
%>