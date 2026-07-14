<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParDeleteEntry.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 20,2010
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
dim sName,sShortName,sAddress1,sAddress2,sPincode,sCity,sState,sCountry,sPhone,sFax
dim sEmail,sWebsite,sECCNo,sSalesLocal,sSalesCentral,sGroupFlag,sGroup
DIM Temp,iCounter,arrUnit,arrUnitName,sPanNo,sMobile,iPartyCode,sSelUnit
Dim oDOM,Root,newElem,newElem1,sCallTy,Temparr,sDelType
dim sQuery,iRecCount,objRs,sHisno,sHisRes,sTinNumber,iUnitCount

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")

iPartyCode = Request.QueryString("hPartyCode")
scallTy = Request.QueryString("hCallTy")
sDelType = Request.QueryString("hDelTy")
sSelUnit = Request.QueryString("hDelTy")
sHisRes = "Deletion "

iUnitCount = 0

Response.Write "sCallTy="& sCallTy
Response.Write "sDelType = "& sDelType

con.begintrans

IF CStr(sCallTy) = "D" Then
	'iPartyCode = Request("optParty")
	Temparr = Split(iPartyCode,"?")
	iPartyCode = Trim(Temparr(3))

	sQuery = "Select isNull(Max(HistoryNo),0) + 1  from APP_M_HistoryPartyMaster "
	With Objrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With

	Set Objrs.ActiveConnection = Nothing
	IF Not Objrs.EOF Then
		sHisNo = Objrs(0)
	End IF
	Objrs.Close

	'1 Marked For Deletion
	'0 Marked For Currently Using it

	sQuery = "Insert Into APP_M_HistoryPartyMaster Select *,"&sHisNo&",'A',"&getUserID&",getDate(),'"&sHisRes&"' "&_
			 "From APP_M_PartyMaster Where PartyCode = "&iPartyCode&" "

	Response.Write sQuery

'	Con.Execute sQuery

	IF CStr(sDelType) = "A" Then
	
		sQuery = "Delete from APP_R_OrgParty  Where PartyCode = "&iPartyCode
		Response.Write sQuery
		con.execute sQuery
		
		sQuery = "Delete from APP_R_AgentOrgParty  Where PartyCode = "&iPartyCode
		Response.Write sQuery
		Con.Execute sQuery
		
		sQuery = "Delete from APP_M_PartyContactPersons where PartyCode = "& iPartyCode
		Response.Write sQuery
		con.execute sQuery

		sQuery = "Delete  from APP_M_PartyLocations where PartyCode = "& iPartyCode
		Response.Write sQuery
		con.execute sQuery
		
		sQuery = "Delete from  ACC_T_PartyOpeningAmt where PartyCode = "& iPartyCode
		Response.Write sQuery
		con.execute sQuery
		
		sQuery = "Delete from APP_M_PartyMaster WHERE PartyCode = "&iPartyCode&" "
		Response.Write sQuery
		Con.Execute sQuery
		
		sQuery = "Insert Into APP_M_HistoryAgentMaster Select *,"&sHisNo&" From  "&_
				 "APP_M_AgentMaster Where AgentCode = "&iPartyCode&" "
		Response.Write sQuery
		Con.Execute sQuery
	

		sQuery = "Delete from APP_M_AgentMaster WHERE AgentCode = "&iPartyCode&" "
		Response.Write sQuery
		Con.Execute sQuery
		
	Else

		IF CStr(sSelUnit) <> "" Then
			
			sQuery = "Select distinct OUDefinitionID from APP_R_OrgParty where PartyCode = "& iPartyCode 
			objRs.Open sQuery,con
			if not objRs.EOF then
				do while not objRs.EOF 
					iUnitCount = iUnitCount + 1
					objRs.MoveNext 
				loop
			end if
			objRs.Close 
			Response.Write "iUnitCount = "& iUnitCount
			
			
			sQuery = "Delete from APP_R_OrgParty  Where PartyCode = "&iPartyCode&" and "&_
					 "OUDefinitionID IN ("&sSelUnit&") "
			Response.Write sQuery
			Con.Execute sQuery
			
			sQuery = "Delete from APP_R_AgentOrgParty  Where PartyCode = "&iPartyCode&" and "&_
					 "OUDefinitionID IN ("&sSelUnit&") "
			Response.Write sQuery
			Con.Execute sQuery
			
			sQuery = "Delete from  ACC_T_PartyOpeningAmt where OUDefinitionID in ("& sSelUnit &") and PartyCode = "& iPartyCode
			Response.Write sQuery
			con.execute sQuery
			
			if iUnitCount = 1 then
				sQuery = "Delete from APP_M_PartyContactPersons where PartyCode = "& iPartyCode
				Response.Write sQuery
				con.execute sQuery

				sQuery = "Delete  from APP_M_PartyLocations where PartyCode = "& iPartyCode
				Response.Write sQuery
				con.execute sQuery
			
			
				sQuery = "Delete from APP_M_PartyMaster WHERE PartyCode = "&iPartyCode&" "
				Response.Write sQuery
				Con.Execute sQuery
			
				sQuery = "Insert Into APP_M_HistoryAgentMaster Select *,"&sHisNo&" From  "&_
					 "APP_M_AgentMaster Where AgentCode = "&iPartyCode&" "
				Response.Write sQuery
				Con.Execute sQuery

				sQuery = "Delete from APP_M_AgentMaster WHERE AgentCode = "&iPartyCode&" "
				Response.Write sQuery
				Con.Execute sQuery
			end if ' if iUnitCount = 1 then
		End IF

	End IF

End IF
'	con.rollbacktrans
'	Response.End 

	Response.Clear 
	con.committrans
%>
