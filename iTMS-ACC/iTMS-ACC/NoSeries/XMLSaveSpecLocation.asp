<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLSaveSpecLocation.asp
	'Module Name				:	Production ( Master)
	'Author Name				:	Kalai selvi 
	'Created On					:	July 31, 2009
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
	Dim oDOM,sMod,sName,sPath,sAddSessionID,sSaveFileName
	
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	
	sName=trim(Request("Name"))
	sMod=trim(Request("Mod"))
	sPath=trim(Request("Path"))
	sAddSessionID = trim(Request("AddSessionID"))
	
	sSaveFileName = sName
	if sMod <> "" then
		sSaveFileName = trim(sSaveFileName) & "_" & sMod
	end if
	if sAddSessionID = "Y" then
		sSaveFileName = trim(sSaveFileName) & "_" & Session.SessionID
	end if  

	sSaveFileName = trim(sSaveFileName) & ".xml"
	
	if trim(sPath) = "" then
		sPath = "../temp/master/"
		sPath = "temp/master/"
	else
		sPath = sPath  & "/"	
	end if 
	
	
	oDOM.async = false
	oDOM.load(Request)
	oDOM.Save server.MapPath(sPath & sSaveFileName)	
	set oDOM=nothing
%>