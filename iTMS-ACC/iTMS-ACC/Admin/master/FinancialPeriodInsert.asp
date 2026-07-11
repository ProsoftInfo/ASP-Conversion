<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	FinancialPeriodInsert.asp
	'Module Name				:	Financial Period
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 13, 2006
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Financial Period</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert("Financial Period has been created / updated successfully");
			window.location.href = "FinancialPeriodEntry.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>

<%
	dim sYear,dcrs,sSql,sFrmYear,arrYear,sToYear,sUserID
	
	sUserId = session("userid")

	set dcrs = server.CreateObject("ADODB.Recordset")

	sYear = Request.Form("txtYear")
	if Request.Form("selPFinStartDate") <> "new" then
		arrYear = split(Request.Form("selPFinStartDate"),":")
		sFrmYear = arrYear(0)
		sToYear = arrYear(1)
	else
		sFrmYear = "01/04/"&sYear
		sToYear = "31/03/"&sYear+1
	end if

	con.beginTrans

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT SUBSTRING(CONVERT(CHAR,FROMPERIOD,103),7,4) FROM MS_FINANCIALPERIOD WHERE SUBSTRING(CONVERT(CHAR,FROMPERIOD,103),7,4) = " & Pack(sYear)
		'Response.Write dcrs.Source
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF  then
		sSql = "UPDATE MS_FINANCIALPERIOD SET ACTIVE = 'Y' WHERE FROMPERIOD= " & Pack(sFrmYear)
		'Response.Write sSql & "<BR>"
		con.Execute sSql
	else
		sSql = "UPDATE MS_FINANCIALPERIOD SET ACTIVE = 'N'"
		'Response.Write sSql & "<BR>"
		con.Execute sSql

		sSql = "INSERT INTO MS_FINANCIALPERIOD(FROMPERIOD,TOPERIOD,ACTIVE,Closed,FinYearCreatedBy,FinYearCreatedOn) VALUES (CONVERT(DATETIME," & Pack(sFrmYear) & ",103),CONVERT(DATETIME," & Pack(sToYear) & ",103),'Y','N',"&sUserID&",getDate())"
		'Response.Write sSql & "<BR>"
		con.Execute sSql
	end if
	dcrs.close

	if con.Errors.count <> 0 then
		dim iCounter
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & "<BR>"
		next
		'Redirect to Error Handling System
	else
		'con.RollbackTrans
		con.CommitTrans
	end if

	con.close
	set con = nothing
%>
<BODY onLoad = "msgbox('','Y')">