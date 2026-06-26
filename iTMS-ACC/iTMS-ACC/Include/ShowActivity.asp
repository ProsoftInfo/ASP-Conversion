<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ShowActivity.asp
	'Module Name				:	Menu
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 08, 2003
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
	dim iSAApplication,iSAProcess,iSAActivity,sSAPath,sApplicationName,iSAActTempNo
	dim sTemp,arrTemp,sTempVal,sTempSp
	
	iSAApplication = Request.QueryString("iApplication")
	iSAProcess = Request.QueryString("iProcess")
	iSAActivity = Request.QueryString("iActivity")
	sSAPath = Request.QueryString("sPath")
	iSAActTempNo=Request.Querystring("iActTempNo")
	
	arrTemp = split(Request.ServerVariables("HTTP_REFERER"),"/")
	sApplicationName = "../"&arrTemp(3)&"/"
	
	Session.Contents.Remove("iApplication")
	Session.Contents.Remove("iProcess")
	Session.Contents.Remove("iActivity")
	Session.Contents.Remove("iActTempNo")

	Session("iApplication") = iSAApplication
	Session("iProcess") = iSAProcess
	Session("iActivity") = iSAActivity
	Session("iActTempNo") = iSAActTempNo

	sTempSp = Split(sSAPath,"/")
	IF Cstr(sTempSP(0)) = "NOSERIES" Then
		Response.Redirect "/"&sSAPath
	End IF
	Response.Write sApplicationName&sSAPath
	
	sSAPath =  replace(sSaPath,":","&")
	Response.Write sApplicationName&sSAPath
	
	Response.Redirect sApplicationName&sSAPath
%>