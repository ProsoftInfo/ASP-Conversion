<%Option Explicit%>
<%
	'Program Name				:	GetOutStandingCredit.asp
	'Module Name				:	Get Out Standing Credit Information for the Particular party unit and Party Type wise
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Sep 14,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
Dim rsTemp
Dim sQuery,sPartyType,sPartySubType,sPartyCode,sPartyUnit,sReturnValue

set rsTemp = Server.CreateObject("ADODB.Recordset")

sPartyType = Request("PartyType")
sPartySubType = Request("PartySubType")
sPartyCode = Request("PartyCode")
sPartyUnit = Request("PartyUnit")

if Trim(sPartyType)<>"" then
   sReturnValue = GetCreditOutStanding(sPartyType,sPartySubType,sPartyCode,sPartyUnit)
   Response.Write sReturnValue
end if

Function GetCreditOutStanding(sPartyType,sPartySubType,sPartyCode,sPartyUnit)
    sQuery = "Select Isnull(Sum(AmountReceivable),0)-Isnull(Sum(AmountReceived),0) From Acc_T_Receivables Where OuDefinitionId='"& sPartyUnit &"' and AmountReceivable>AmountReceived and partycode=" & sPartyCode &" and partysubtype=" & sPartySubType &" and PartyType='"& sPartyType & "'"
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        GetCreditOutStanding = rsTemp(0)
    end if
    rsTemp.Close 
End Function
%>

