<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLSave.asp
	'Module Name				:	Inventory (Common Program)
	'Author Name				:	TAJUDEEN
	'Created On					:	June 11, 2004
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
	
	bFlag = Request.QueryString("SessionFlag")
	sValue = Request.QueryString("Value")
	sFolder = Request.QueryString("Folder")
	
	if bFlag then
		NewXml.save server.MapPath("..\temp\" & sFolder & "\" & sValue & Session.SessionID & ".xml")
	else
		NewXml.save server.MapPath("..\temp\" & sFolder & "\" & sValue & ".xml")
	end if
	
%>
