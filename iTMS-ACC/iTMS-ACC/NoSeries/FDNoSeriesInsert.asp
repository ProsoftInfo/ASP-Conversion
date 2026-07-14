<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	FDNoSeriesInsert.asp
	'Module Name				:	Sales (Master Creation)
	'Author Name				:	Subbiah
	'Created On					:	Sep 03,2003
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->

<%
dim iUnitNo,iActivity,iCounter,iInvSeriesCode,sType
dim sSql,iExistBookNo,sItmType,sActName
dim iSeries,iSeriesType,bPayRecNo,iLength,sAgentcode

'iUnitNo=trim(Request.Form("selUnit"))
iUnitNo = Session("organizationcode")
sItmType=trim(Request.Form("hoptval"))
iActivity=trim(Request.Form("selActType"))
iSeries=trim(Request.Form("selNoSeries"))
iSeriesType=trim(Request.Form("hSeriesType"))
iLength=trim(Request.Form("hSeriesLen"))
sActName = trim(Request.Form("hActivityName"))
'sAgentcode = trim(Request.Form("selAgent"))
con.BeginTrans
'if Trim(sItmtype) = "" then
	'sItmtype = 0
'End if
'if Trim(sAgentcode) = "" then
	'sAgentcode = 0
'End if
'Response.Write "<p>Unit="&iUnitNo & " " & iSeries
'Response.Write "<p>iInvSeriesCode="&iInvSeriesCode
' Module Lot Number
'sOrgid,iAppcode,iModuleCode,iSeriesNo,sType,sName,sDescription,iLen
iInvSeriesCode=GenSeriesCode(iUnitNo,"8","3",iSeries,iSeriesType,"",sActName,iLength)
'Response.Write "<p>iInvSeriesCode="&iInvSeriesCode

sSql = "INSERT INTO FDP_M_Noseries (ORGANISATIONCODE,DEPOSITTYPE," &_
	"ACTIVITYTYPE,SERIESNO,SERIESCODE) VALUES " &_
	"('" & iUnitNo & "','" & sItmType & "','"&iActivity& "',"& iSeries & "," & iInvSeriesCode & ")"
Response.Write sSql
con.Execute sSql

if con.Errors.count <> 0 then
	dim iErrCounter
	con.RollbackTrans
	for iErrCounter=0 to con.Errors.count
		Response.Write con.Errors(iErrCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
Response.Redirect "FDNoSeriesEntry.asp"
%>      
