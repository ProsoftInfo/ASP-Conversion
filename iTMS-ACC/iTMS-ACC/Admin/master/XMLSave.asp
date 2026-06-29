<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLSave.asp
	'Module Name				:	Admin(Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jun 30,2012
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
	'This program Saves the XML file with the given Name 
	'with\without SessionID in the given Directory
	
	'Call the Program like XMLSave.asp?SessionFlag=True&Value=TempXML&Folder=Master
	dim NewXml, sValue, bFlag, sFolder

	Set NewXml = Server.CreateObject("Microsoft.XMLDOM")
	
	NewXml.async = false
	NewXml.load(Request)
	
	sValue = Request.QueryString("Name")
	
	NewXml.save server.MapPath("../temp/" & sValue &"_"& Session.SessionID & ".xml")
	set NewXml=nothing
%>
