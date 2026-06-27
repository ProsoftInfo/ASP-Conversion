<%
	'Program Name				:	GetAccHeadForPartyType.asp
	'Module Name				:	Get Account Head Information for the Party Type
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Oct 03,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:
%>

<%
Dim rsTemp
Dim sPartyType,sPartySubType,sPartyCode,sPartyUnit,sReturnValue
Dim sCreditLimit,sCreditDays

set rsTemp = Server.CreateObject("ADODB.Recordset")

Function GetAccHeadForPartyType(sPartyType,sPartySubType,sPartyCode,sPartyUnit)
    sQuery = " Select AccountHead from ACC_R_OrgPartyType where PartyType = '"& sPartyType &"' and PartySubType="& sPartySubType &" and OUDefinitionID = '"& sPartyUnit &"'"
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        sReturnValue = rsTemp(0)
    else
        sReturnValue = "0"
    end if
    rsTemp.Close 
    GetAccHeadForPartyType = sReturnValue
End Function
%>

