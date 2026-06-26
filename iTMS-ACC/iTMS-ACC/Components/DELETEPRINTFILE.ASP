<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	DeletePrintFile.asp
	'Module Name				:	
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 30, 2004
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
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
	dim sFilePath,objfs
	sFilePath = trim(Request("sPath"))
	
	Set objfs = CreateObject("Scripting.FileSystemObject")
	
	if objfs.FileExists(server.MapPath(sFilePath)) then
		objfs.DeleteFile server.MapPath(sFilePath)
	end if
%>
