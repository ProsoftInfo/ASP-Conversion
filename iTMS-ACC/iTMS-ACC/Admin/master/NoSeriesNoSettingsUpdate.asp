<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	NoSeriesNoSettingsUpdate.asp
	'Module Name				:	ADMIN (MASTER)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	FEB 19,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Work Center</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	Dim sOrgCode, sPackNum, iSeriesNo, iSeriesCode, sPeriod, sSql
	Dim oDOM,nRoot,nSubNode
	Dim sTempModule,stempActivity,sItemType,sClassCode,sCatCode
	
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	oDOM.async = false
	oDOM.load(server.MapPath("../temp/NoSeriesSettings_"&Session.SessionID&".xml"))
	'Response.Write oDOM.xml	
	Response.Write "<font color=red>"
	sTempModule = Request.QueryString("Module")
	stempActivity = Request.QueryString("Activity")
	sItemType= Request.QueryString("Item")
	sClassCode = Request.QueryString("ClassCode")
	sCatCode = Request.QueryString("CatCode")

	set	nRoot = oDOM.documentElement
	sOrgCode = nRoot.getAttribute("OrgCode")
	Con.Begintrans
	for each nSubNode in nRoot.childNodes
		
		iSeriesNo = nSubNode.getAttribute("SeriesNo")
		iSeriesCode = nSubNode.getAttribute("SeriesCode")
		sPeriod = nSubNode.getAttribute("Period")
		sPackNum = nSubNode.getAttribute("NewPack")
		
		if trim(sPackNum)<>"" then
			sSql = "UPDATE APP_R_NOSERIESMODULEENTRY SET NUMBER = "&sPackNum&" WHERE OUDEFINITIONID = '"&sOrgCode&"' AND SERIESNO = "&iSeriesNo&" AND SERIESCODE = "&iSeriesCode&" AND PERIOD = '"&sPeriod&"'"
			Response.Write sSql
			con.execute sSql
		end if 'if trim(sPackNum)<>"" then
	next
	
		Response.Write "<p>sTempModule ="&sTempModule
	Response.Write "<p>stempActivity = "&stempActivity
	Response.Write "<p>sItemType= "& sItemType 
	Response.Write "<p>CatCode = "& sCatCode 
	Response.Write "<p>ClassCode = "& sClassCode 

	
	

	If Con.Errors.count <> 0 Then
		Con.RollbackTrans
		For iCounter = 0 to Con.Errors.count
			Response.Write Con.Errors(iCounter) &"<br>"
		Next
	Else
	   ' Con.RollbackTrans
	    'Response.End 
	    Response.Clear 
		Con.CommitTrans
	End If
	Con.close
	Set Con = Nothing	
	Response.Redirect ("../../admin/Master/NoSeriesNoSettings.asp?hSelModule="&sTempModule &"&hSelActivity="&stempActivity&"&hCatCode="&sCatCode&"&hClassCode="&sClassCode)
%>
