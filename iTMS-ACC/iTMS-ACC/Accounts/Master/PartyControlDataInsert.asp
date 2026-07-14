<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PartyControlDataInsert.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 13,2011
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
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
    Dim oDOM,ndRoot,ndChild
    Dim sUnitId,sPartyType,sSubType,sPartyCode,sQuery
    Dim iCreditLimit,iCreditDays
    
    set oDOM = Server.CreateObject("Microsoft.XMLDOM")
    oDOM.load(Request)
    
    con.begintrans
    set ndRoot = oDOM.documentElement
    
    sPartyCode = ndRoot.getAttribute("PartyCode")
    if ndRoot.hasChildNodes() then
        for each ndChild in ndRoot.childNodes
            if ndChild.getAttribute("Code")=sPartyCode then
                sUnitId      = ndChild.getAttribute("Unit")
                sPartyType   = ndChild.getAttribute("Type")
                sSubType     = ndChild.getAttribute("SubType")
                iCreditLimit = ndChild.getAttribute("CreditLimit")
                iCreditDays  = ndChild.getAttribute("CreditDays")
                
                if iCreditLimit = "" then iCreditLimit = "NULL"
                if iCreditDays = "" then iCreditDays = "NULL"
                
                sQuery = "Update APP_R_OrgParty set CreditLimit ="& iCreditLimit &",CreditDays = "& iCreditDays &" where PartyType='"& sPartyType &"' and PartySubType="& sSubType &" and OUDefinitionID = '"& sUnitId &"' and PartyCode = "& sPartyCode 
                Response.Write vbCrLf &vbCrLf & sQuery & vbCrLf &vbCrLf
                con.execute sQuery
            end if 'if ndChild.getAttribute("Code")=sPartyCode then
        next
    end if
    
'	con.rollbacktrans
'	Response.End 

	Response.Clear 
	con.committrans
%>
