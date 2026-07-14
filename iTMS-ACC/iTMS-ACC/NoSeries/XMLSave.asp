<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLSave.asp
	'Module Name				:	Sales (Master)
	'Author Name				:	Manohar Prabhu
	'Created On					:	June 07, 2004
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
	Dim oDOM,sMod,sName
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	sMod=Request("Mod")
	sName=Request("Name")

	oDOM.async = false
	oDOM.load(Request)
	oDOM.Save server.MapPath("./temp/master/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")
	set oDOM=nothing
%>