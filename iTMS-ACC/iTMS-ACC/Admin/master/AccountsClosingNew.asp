<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccountsClosingNew.asp
	'Module Name				:	Admin - Accounts Transfer
	'Author Name				:	TAJUDEEN S
	'Created On					:	June 07, 2004
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag,sFor) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "CloseEntry.asp?Frm="+sFor;
		}
		else {
			alert(strr);
			window.location.href = "CloseEntry.asp?Frm="+sFor;
		}
	}
//-->
</SCRIPT>

<%

	'XML DOM Variables
	Dim adoCmd,sWho,sSql,sfrmFinMonYr,stoFinMonYr,sNewFinFYr,sNewFinTYr,sFor,arrTemp,sCurrDate
	Dim dPreFinStartDate,dPreFinEndDate,sUnitCode,iUserID,iUnitCnt,iAudCnt,Objrs

	Set adoCmd = Server.CreateObject("ADODB.Command")
	Set Objrs = Server.CreateObject("ADODB.RECORDSET")


	Set adoCmd.ActiveConnection = con
	Response.Write Request.QueryString

	sWho = Request.QueryString("sWho")
	sfrmFinMonYr = Request("CurrFromDate")
	stoFinMonYr = Request("CurrToDate")
	sFor = Request("hFor")
	sUnitCode = Request("UnitCode")
	iUserID = session("userid")

	'arrTemp = split(trim(Request.Form("selPFinStartDate")),":")
	'Response.Write "hPFinStartDate="& Request.Form("selPFinStartDate")
	dPreFinStartDate = Request("PrevFromDate")
	dPreFinEndDate = Request("PrevToDate")

'	Response.Write "<Br>sUnitCode = "& sUnitCode
'	Response.Write "<Br>sfrmFinMonYr ="& sfrmFinMonYr
'	Response.Write "<Br>stoFinMonYr ="& stoFinMonYr
'	Response.Write "<br>dPreFinStartDate = "& dPreFinStartDate
'	Response.Write "<br>dPreFinEndDate = "& dPreFinEndDate
'	Response.Write "<br>sFor  = "& sFor

	''Response.Write dPreFinStartDate
	'Response.End

	'sNewFinFYr = Request.Form("hPFinStartDate")
	'sNewFinTYr = Request.Form("hPFinEndDate")

	sNewFinFYr = dPreFinStartDate
	sNewFinTYr = dPreFinEndDate

	sNewFinFYr = mid(sNewFinFYr,7,4)&mid(sNewFinFYr,4,2)
	sNewFinTYr = mid(sNewFinTYr,7,4)&mid(sNewFinTYr,4,2)

	sfrmFinMonYr = mid(sfrmFinMonYr,4,2)&mid(sfrmFinMonYr,7,4)
	stoFinMonYr = mid(stoFinMonYr,4,2)&mid(stoFinMonYr,7,4)

	sCurrDate = Cstr(Date())

	'Response.End
	'Response.Write sNewFinFYr & " -- " & sNewFinTYr & "<BR>"
	'Response.Write sfrmFinMonYr & " -- " & stoFinMonYr

	''Response.Write sFor
	'Response.End

	'Response.write sNewFinFYr &" " & sNewFinTYr &" " & sfrmFinMonYr &" "& stoFinMonYr
	'Response.end

	con.begintrans
	if CStr(sFor) = "GL" then 'For GL Account Head Transfer then
		' CALLING THE STORED PROCEDURE
		sSql = "GLYearTrans"
		adoCmd.CommandText = sSql
		adoCmd.CommandType = 4 'adCmdStoredProc
		adoCmd.Parameters.Append adoCmd.CreateParameter("@frmFinMonYr",3,1,6,sNewFinFYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@toFinMonYr",3,1,6,sNewFinTYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinFYr",129,1,6,sfrmFinMonYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinTYr",129,1,6,stoFinMonYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@TransferredBy",3,1,6,iUserID)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@TransferredOn",129,1,10,sCurrDate)
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@ForUnit",200,1,Len(sUnitCode),sUnitCode)
		adoCmd.Execute()

		Set adoCmd = Nothing
		Set adoCmd = Server.CreateObject("ADODB.Command")
		Set adoCmd.ActiveConnection = con

		sSql = "PartyYearTrans"
		adoCmd.CommandText = sSql
		adoCmd.CommandType = 4 'adCmdStoredProc
		adoCmd.Parameters.Append adoCmd.CreateParameter("@frmFinMonYr",3,1,6,sNewFinFYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@toFinMonYr",3,1,6,sNewFinTYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinFYr",129,1,6,sfrmFinMonYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinTYr",129,1,6,stoFinMonYr)
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@ForUnit",200,1,Len(sUnitCode),sUnitCode)
		adoCmd.Execute()
		Set adoCmd = Nothing

		'sSql =  "UPDATE Ms_FinancialPeriod SET FinYearClosedBy = "&iUserID&",  "&_
		'		"FinYearClosedOn = getDate() WHERE Convert(Varchar,FromPeriod,103) = '"&dPreFinStartDate&"' "
		'Con.Execute sSql

	elseif CStr(sFor) = "AG" then 'Audit GL
		' CALLING THE STORED PROCEDURE
		sSql = "AuditGLYearTrans"
		adoCmd.CommandText = sSql
		adoCmd.CommandType = 4 'adCmdStoredProc

		adoCmd.Parameters.Append adoCmd.CreateParameter("@frmFinMonYr",3,1,6,sNewFinFYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@toFinMonYr",3,1,6,sNewFinTYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinFYr",129,1,6,sfrmFinMonYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinTYr",129,1,6,stoFinMonYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@ForUnit",200,1,Len(sUnitCode),sUnitCode)
		adoCmd.Execute()
		Set adoCmd = Nothing
		Set adoCmd = Server.CreateObject("ADODB.Command")
		Set adoCmd.ActiveConnection = con

		sSql = "AuditPartyYearTrans"
		adoCmd.CommandText = sSql
		adoCmd.CommandType = 4 'adCmdStoredProc
		adoCmd.Parameters.Append adoCmd.CreateParameter("@frmFinMonYr",3,1,6,sNewFinFYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@toFinMonYr",3,1,6,sNewFinTYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinFYr",129,1,6,sfrmFinMonYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinTYr",129,1,6,stoFinMonYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@ForUnit",200,1,Len(sUnitCode),sUnitCode)
		adoCmd.Execute()
		Set adoCmd = Nothing

		sSql =  "INSERT INTO Ms_AuditorClosing (OUDefinitionID, FinYearAuditClosedBy, "&_
				"FinYearAuditClosedOn,FromPeriod, ToPeriod) "&_
				"VALUES ('"&sUnitCode&"', "&iUserID&", getdate(), Convert(datetime,'"&dPreFinStartDate&"',103),  "&_
				"Convert(datetime,'"&dPreFinEndDate&"',103)) "

		Con.Execute sSql

		sSql =  "UPDATE Ms_FinancialPeriod SET FinYearClosedBy = "&iUserID&",  "&_
				"FinYearClosedOn = getDate() WHERE Convert(Datetime,FromPeriod,103) = Convert(datetime,'"&dPreFinStartDate&"',103)"
		Con.Execute sSql

		'sSql = "SELECT COUNT(1) FROM DCS_OrganizationUnitDefinitions WHERE LEN(OUDefinitionID) > 4 "
		'Objrs.Open sSql,Con
		'IF Not Objrs.EOF Then
		'	iUnitCnt = Objrs(0)
		'Else
		'	iUnitCnt = 0
		'End IF
		'Objrs.Close

		'sSql = "SELECT COUNT(1) FROM Ms_AuditorClosing WHERE Convert(datetime,FromPeriod,103) = Convert(datetime,'"&dPreFinStartDate&"',103)"
		'Objrs.Open sSql,Con
		'IF Not Objrs.EOF Then
		'	iAudCnt = Objrs(0)
		'Else
		'	iAudCnt = 0
		'End IF
		'Objrs.Close

		'IF CStr(iUnitCnt) = CStr(iAudCnt) Then
		'	sSql =  "UPDATE Ms_FinancialPeriod SET Closed = 'Y'  "&_
		'			"WHERE Convert(datetime,FromPeriod,103) = Convert(datetime,'"&dPreFinStartDate&"',103)"
		'	Con.Execute sSql
		'End IF
	Elseif CStr(sFor) = "PROLOSS" then 'Update GL,Party and Profit and Loss Account
		Dim sRetVal
		sSql = "GLPartyPRLossUpdate"
		adoCmd.CommandText = sSql
		adoCmd.CommandType = 4 'adCmdStoredProc
		adoCmd.Parameters.Append adoCmd.CreateParameter("@frmFinMonYr",3,1,6,sNewFinFYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@toFinMonYr",3,1,6,sNewFinTYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinFYr",129,1,6,sfrmFinMonYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewFinTYr",129,1,6,stoFinMonYr)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@ForUnit",200,1,Len(sUnitCode),sUnitCode)
		adoCmd.Parameters.Append adoCmd.CreateParameter("RetVal",3,4)
		adoCmd.Execute()
		sRetVal = adoCmd.Parameters("RetVal")
		Set adoCmd = Nothing



	end if


	''Response.Clear
	If con.Errors.count <> 0 Then
		Dim iErrCounter
		con.RollbackTrans
		For iErrCounter=0 to con.Errors.count - 1
			Response.Write con.Errors(iErrCounter) & "<BR>"
		Next
		'Redirect to Error Handling System
	Else
		'con.RollbackTrans
		con.CommitTrans
		'Response.End
		if CStr(sFor) = "GL" then
%>

	<BODY onLoad = "msgbox('Closing Amount for GL and Party Account Balance has been Transferred Successfully','Y','<%=sFor%>')">
<%
		elseif CStr(sFor) = "PA" then
%>
	<BODY onLoad = "msgbox('Closing Amount for Party control Account has been Transferred Successfully','Y','<%=sFor%>')">

<%
		elseif CStr(sFor) = "AG" then
%>
	<BODY onLoad = "msgbox('Auditor Closing for GL and Party Account balance has been Transferred Successfully','Y','<%=sFor%>')">

<%
		elseif CStr(sFor) = "AP" then
%>
	<BODY onLoad = "msgbox('Auditor Closing for Party Balance has been Transferred Successfully','Y','<%=sFor%>')">

<%
		end if
	end if
	con.close
	Set con = Nothing
%>
