<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLGetPayables.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 19,2003
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<%
dim sOrgId,sPartyValue,objRs,sQuery,sParCode,sParSubType,sParType
dim OutData,newElem,Root
dim iDocNo,iSno
dim sAmtPaid,sAmtPayable,sInvDate,sInvNo,sVouDate,sVouNo

sOrgId=Request("orgid")
sPartyValue=split(trim(Request("ParCode")),"?")

sParType=sPartyValue(0)
sParSubType=sPartyValue(1)
sParCode=sPartyValue(3)

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set OutData = Server.CreateObject("Microsoft.XMLDOM")
						
sQuery ="select PayablesNumber,TransactionNumber,VoucherDate,isnull(PartyBillNumber,0),isnull(PartyBillDate,''), "&_
	"AmountPayable,AmountPaid from Acc_T_Payables where OUDefinitionID='"&sOrgId&"' and PartyType='"&sParType&"'"&_
	" and PartySubType="&sParSubType&" and PartyCode="&sParCode&" and AmountPayable>AmountPaid"

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

set iDocNo = objRs(0)
set sVouNo = objRs(1)
set sVouDate = objRs(2)
set sInvNo = objRs(3)
set sInvDate = objRs(4)
set sAmtPayable = objRs(5)
set sAmtPaid = objRs(6)

iSno=1
If not objRs.EOF then
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root
	
	Do While Not objRs.EOF
			Set newElem = OutData.createElement("Doc")
			newElem.setAttribute "No", iDocNo
			newElem.setAttribute "VouNo",sVouNo
			newElem.setAttribute "VouDate",sVouDate
			newElem.setAttribute "InvNo",sInvNo 
			newElem.setAttribute "InvDate",sInvDate
			newElem.setAttribute "AmtPayable",sAmtPayable
			newElem.setAttribute "AmtPaid",sAmtPaid
			newElem.setAttribute "AmtAdjust","0"
			Root.appendChild newElem
			objRs.MoveNext
	loop
	Response.ContentType="text/xml"
	Response.Write OutData.xml
end if
objRs.Close
%>                                            
