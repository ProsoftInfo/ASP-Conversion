<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	InventoryNoSeriesInsert.asp
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	August 01,2003
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
Dim iUnitNo,iActivity,iCounter,iInvSeriesCode,sType
Dim sSql,iExistBookNo,sItmType,sActName
Dim iSeries,iSeriesType,bPayRecNo,iLength

iUnitNo=trim(Request.Form("selUnit"))
sItmType=trim(Request.Form("selItmType"))
iActivity=trim(Request.Form("selActType"))
iSeries=trim(Request.Form("selNoSeries"))
iSeriesType=trim(Request.Form("hSeriesType"))
iLength=trim(Request.Form("hSeriesLen"))
sActName = trim(Request.Form("hActivityName"))

con.BeginTrans

'sOrgid,iAppcode,iModuleCode,iSeriesNo,sType,sName,sDescription,iLen
' Module Lot Number
if iActivity = "LO" then
	iInvSeriesCode=GenSeriesCode(iUnitNo,"4","3",iSeries,iSeriesType,"",sActName,iLength)
' Module Packing Number
elseif iActivity = "PN" then
	iInvSeriesCode=GenSeriesCode(iUnitNo,"4","4",iSeries,iSeriesType,"",sActName,iLength)
' Module Sample Label
elseif iActivity = "SL" then
	iInvSeriesCode=GenSeriesCode(iUnitNo,"2","3",iSeries,iSeriesType,"",sActName,iLength)
' Module DC - Gate Pass
elseif iActivity = "DC" then
	iInvSeriesCode=GenSeriesCode(iUnitNo,"4","5",iSeries,iSeriesType,"",sActName,iLength)
elseif iActivity = "IS" then
	iInvSeriesCode=GenSeriesCode(iUnitNo,"4","1",iSeries,iSeriesType,"",sActName,iLength)
elseif iActivity = "MR" then
	iInvSeriesCode=GenSeriesCode(iUnitNo,"4","1",iSeries,iSeriesType,"",sActName,iLength)
end if

'sSql = "INSERT INTO INV_M_NUMBERSERIES (ORGANISATIONCODE,ITEMTYPE," &_
'	"ACTIVITYTYPE,SERIESNO,SERIESCODE) VALUES " &_
'	"('" & iUnitNo & "','" & sItmType & "','"&iActivity& "',"& iSeries & "," & iInvSeriesCode & ")"

sSql = "INSERT INTO INV_M_NUMBERSERIES (ORGANISATIONCODE," &_
	"ACTIVITYTYPE,SERIESNO,SERIESCODE) VALUES " &_
	"('" & iUnitNo & "','"&iActivity& "',"& iSeries & "," & iInvSeriesCode & ")"
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
Response.Redirect "InventoryNoSeriesEntry.asp"
%>