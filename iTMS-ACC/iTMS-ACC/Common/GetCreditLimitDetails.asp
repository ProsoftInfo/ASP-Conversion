<%Option Explicit%>
<%
	'Program Name				:	GetCreditLimitDetails.asp
	'Module Name				:	Get Credit Limit Information for the Particular party unit and Party Type wise
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Sep 13,2011
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
Dim sCreditLimit,sCreditDays

set rsTemp = Server.CreateObject("ADODB.Recordset")

sPartyType = Request("PartyType")
sPartySubType = Request("PartySubType")
sPartyCode = Request("PartyCode")
sPartyUnit = Request("PartyUnit")

if Trim(sPartyType)<>"" then
   sReturnValue = GetCreditLimit(sPartyType,sPartySubType,sPartyCode,sPartyUnit)
   Response.Write sReturnValue
end if

Function GetCreditLimit(sPartyType,sPartySubType,sPartyCode,sPartyUnit)
    sQuery = "Select CreditLimit,CreditDays from APP_R_OrgParty where OUDefinitionID="& sPartyUnit &" and PartyType = '"& sPartyType &"' and PartySubType ="& sPartySubType &" and PartyCode ="& sPartyCode 
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        sCreditLimit = rsTemp(0)
        sCreditDays  = rsTemp(1)
        
        if Trim(sCreditDays)="" or IsNull(sCreditDays) then sCreditDays = ""
        if Trim(sCreditLimit)="" or IsNull(sCreditLimit) then sCreditLimit = ""
        
        GetCreditLimit = sCreditLimit &":"& sCreditDays 
    end if
    rsTemp.Close 
End Function
%>

