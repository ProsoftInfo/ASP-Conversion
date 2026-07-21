<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLSavePur.asp
	'Module Name				:	Purchase(Receipts)
	'Author Name				:	SRIDEVI PRIYA A
	'Created On					:	March 14,2003
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
	
	oDOM.Save server.MapPath("../temp/transaction/"&sName&".xml")	
	
	
	set oDOM=nothing
%>