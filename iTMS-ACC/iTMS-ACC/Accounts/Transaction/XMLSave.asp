<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLSave.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February  13, 2003
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
	If Not oDOM.load(Request) Then
		Response.ContentType = "text/plain"
		Response.Write "Unable to save XML data."
		If oDOM.parseError.errorCode <> 0 Then
			Response.Write " " & Server.HTMLEncode(oDOM.parseError.reason)
		End If
		Response.End
	End If
	oDOM.Save server.MapPath("../temp/transaction/"&sName&"_"&sMod&"_"&Session.SessionID&".xml")	
	set oDOM=nothing
%>
