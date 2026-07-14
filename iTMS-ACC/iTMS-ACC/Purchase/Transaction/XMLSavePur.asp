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
	sDesgDir = trim(Request("ToDir"))
	'Response.Write "<p> sDesgDir = " & sDesgDir 
	
	oDOM.async = false
	oDOM.load(Request)
	if Trim(sDesgDir) = ""  then
		oDOM.Save server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")	
	else
		oDOM.Save server.MapPath("../"& sDesgDir &"/"&sName&".xml")	
	end if 	
	
	set oDOM=nothing
%>