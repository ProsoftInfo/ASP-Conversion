
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SendingEmail.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	Septemper 27,2011
	'Modified By                :   
	'Modified On                :   
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
	Dim sch,cdoConfig,cdoMessage,objBP
	
	Dim sSMTPid,sEmailAdd1,sFrom,sSubject,sMessage
	
	Dim oDOM,Root,TempNode
	
	set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	oDOM.load server.MapPath("../xmldata/EmailSending.xml")

	set Root = oDOM.DocumentElement

	for each TempNode in Root.ChildNodes
		
		if TempNode.NodeName="Email" then
			sSMTPid		= TempNode.getAttribute("SMTPid")
			sEmailAdd1	= TempNode.getAttribute("EmailAdd1")
			sFrom		= TempNode.getAttribute("From")
			sSubject	= TempNode.getAttribute("Subject")
			sMessage	= TempNode.getAttribute("Message")
		end if 	
	next
	
	sch = "http://schemas.microsoft.com/cdo/configuration/"
	 
	Set cdoConfig = Server.CreateObject("CDO.Configuration")
	cdoConfig.Fields.Item(sch & "sendusing")	= 2
	cdoConfig.Fields.Item(sch & "smtpserver")	= sSMTPid
	cdoConfig.fields.update

	Set cdoMessage = Server.CreateObject("CDO.Message")
	Set cdoMessage.Configuration = cdoConfig

	Dim tempArray, iCount, sTempEmail
	
	tempArray = split(sEmailAdd1,";")
	iCount = 0
	sTempEmail = ""
	  
	cdoMessage.From = sFrom
	cdoMessage.Subject = sSubject
	'cdoMessage.TextBody = sMessage
	cdomessage.HTMLBody = sMessage
	
	
	For iCount = 0 To Ubound(tempArray)
		'cdoMessage.To = trim(tempArray(iCount)) 
		cdoMessage.bcc = trim(tempArray(iCount)) 
		'cdoMessage.Send
	Next
	
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
%>	  
	  