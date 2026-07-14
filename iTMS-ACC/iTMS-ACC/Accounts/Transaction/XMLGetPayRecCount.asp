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
	'Program Name				:	XMLGetPayRecCount.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	April 28, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	VouCAEntry.asp
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
<!--#include virtual="/include/Accpopulate.asp"-->
<%
dim objRs,objRs1,sQuery,OutData,Root,newElem,iParTyCode,sRetVal,iAdvCount
dim sorgID,iParSubType,sParType,sParTypeName,iPayCount,iRecCount,sVouDate


set objRs = Server.CreateObject("ADODB.Recordset")


sorgID = Request("orgID")
iParSubType = Request("ParSubType")
sParType = Request("ParType")
iParTyCode = Request("PartyCode")
sVouDate = Request("VouDate")

if trim(sVouDate)="" then
	sVouDate=Day(date)&"/"&Month(date)&"/"&Year(date)
end if


sQuery = "Select Count(AmountPayable) From Acc_T_CreatedPayables Where AmountPayable>AmountPaid "&_
		 "and OUDefinitionID = '"&sorgID&"' and PartyType = '"&sParType&"' and PartySubType = "&iParSubType&" "&_
		 "and PartyCode = "&iParTyCode&" and VoucherDate <= convert(datetime,'"&sVouDate&"',103) "

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
End with

Set objRs.Activeconnection = nothing
IF Not objRs.EOF Then
	iPayCount = objRs(0)
End IF
objRs.Close

sQuery = "select count(AmountReceivable) from Acc_T_CreatedReceivables Where AmountReceivable>AmountReceived "&_
		 "and OUDefinitionID='"&sorgID&"' and PartyType='"&sParType&"' and PartySubType = "&iParSubType&" "&_
		 "and PartyCode = "&iParTyCode&" and VoucherDate <= convert(datetime,'"&sVouDate&"',103) "

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
End with

Set objRs.Activeconnection = nothing
IF Not objRs.EOF Then
	iRecCount = objRs(0)
End IF
objRs.Close

sQuery = "select count(1) from Acc_T_CreatedAdvances Where isNull(AdvancePaid,AdvanceReceived)>isNull(AdvanceAdjusted,0) "&_
		 "and OUDefinitionID='"&sorgID&"' and PartyType='"&sParType&"' and PartySubType = "&iParSubType&" "&_
		 "and PartyCode = "&iParTyCode&" "
		 

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
End with

Set objRs.Activeconnection = nothing
IF Not objRs.EOF Then
	iAdvCount = objRs(0)
End IF
objRs.Close




sRetVal = iPayCount&":"&iRecCount&":"&iAdvCount

Response.Write sRetVal

set objRs=nothing


%>

