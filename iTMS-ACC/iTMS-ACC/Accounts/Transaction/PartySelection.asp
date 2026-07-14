<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PPartySelection.asp
	'Module Name				:	ACCOUNTS
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	May 02, 2004
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
<!--#include virtual="/include/clsDatabase.asp"-->
<!--#include virtual="/include/clsDatalist.asp"-->
<!--#include virtual="/include/populate.asp"-->
<HTML><HEAD><TITLE>Party Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sQuery,sTemp
dim sIType,sOrgID,sFilter,sParType,sParSubType,sParTypeName,sSearchBy,sFinMonYear,sCurrDate
Dim sCurrDay,sCurrMon
sFinMonYear = "04"
sFinMonYear = sFinMonYear&Trim(Left(Session("FinPeriod"),4))
sCurrDay = Day(Now)
sCurrMon = Month(Now)

IF Len(Trim(sCurrDay)) = 1 Then
	sCurrDay = "0"&Trim(sCurrDay)
End IF

IF Len(Trim(sCurrMon)) = 1 Then
	sCurrMon = "0"&Trim(sCurrMon)
End IF

sCurrDate = Trim(sCurrDay&"/"&sCurrMon&"/"&Year(Now))


sOrgID = trim(Request.QueryString("orgID"))
sTemp=split(trim(Request("Party")),"?")
sSearchBy = trim(Request.QueryString("SearchBy"))


sParType=sTemp(0)
sParSubType=sTemp(1)
sParTypeName=sTemp(2)

'Response.write sSearchBy

'Response.Write sParType&" " & sParSubType &"<br>"

sFilter = trim(Request.QueryString("Query"))&"%"

'Response.Write "sSearchBy = "& sSearchBy &"<br>"
'Response.Write "sFilter = "& sFilter &"<br>"


Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist

oDataList.PageSize = 15

oDataList.PrimaryKey = "PartyCode"

IF CStr(sParSubType) = "0" or CStr(sParSubType) = "000" Then

	oDataList.AddDisplayField "Party Name"

	oDataList.AddReturnedField "0"
	oDataList.AddReturnedField "1"
	oDataList.AddReturnedField "2"
	oDataList.AddReturnedField "3"
	oDataList.AddReturnedField "4"

	oDataList.AddReturnedField "0"
	oDataList.AddReturnedField "1"
	oDataList.AddReturnedField "2"
	oDataList.AddReturnedField "1"
	oDataList.AddReturnedField "1"
Else
	oDataList.AddDisplayField "Party Name"
	oDataList.AddDisplayField "Balance"

	oDataList.AddReturnedField "0"
	oDataList.AddReturnedField "2"
	oDataList.AddReturnedField "3"
	oDataList.AddReturnedField "4"
	oDataList.AddReturnedField "5"

	oDataList.AddReturnedField "0"
	oDataList.AddReturnedField "2"
	oDataList.AddReturnedField "3"
	oDataList.AddReturnedField "2"
	oDataList.AddReturnedField "2"
End IF

oDataList.SearchForDesc "Search For"
oDataList.AddSearchField "Party Name < Starts with >","IN"
oDataList.AddSearchField "Party Name < Anywhere >","IA"

IF CStr(sSearchBy) = "" Then
	sSearchBy = "IN"
End IF

IF instr(1,Cstr(sSearchBy),"IN") > 0  Then
	IF CStr(sParSubType) = "0" Then
		if sFilter = "%" then
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"'and PartyType='"&sParType&"' order by PartyName"
		else
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' and PartyType='"&sParType&"'  "&_
							 "and PartyName Like '"&sFilter&"' order by PartyName"
		End IF
	Elseif CStr(sParType) = "CRDR" and Cstr(sParSubType) = "000" Then
		if sFilter = "%" then
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' order by PartyName"
		else
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"'   "&_
							 "and PartyName Like '"&sFilter&"' order by PartyName"
		End IF
	Else
		if sFilter = "%" then
			'oDataList.sSQL = "select PartyName,PartyCode,OrgnPartyCode,PartySubType,PartyType from VwOrgParty "&_
			'				 "where OUDefinitionID='"&sOrgId&"'and PartyType='"&sParType&"' and PartySubType="&sParSubType& " order by PartyName"

			oDataList.sSQL = "SELECT V.PARTYNAME, "&_
							 "ISNULL(DBO.GETPARTYOPEN('"&sFinMonYear&"','"&sCurrDate&"','"&sOrgID&"','"&sParType&"',"&sParSubType&",V.PARTYCODE),'0.00') BALANCE, "&_
							 "V.PARTYCODE,V.ORGNPARTYCODE, V.PARTYSUBTYPE,V.PARTYTYPE FROM VWORGPARTY V WHERE  "&_
							 "V.OUDEFINITIONID = '"&sOrgID&"' AND PARTYTYPE = '"&sParType&"' AND PARTYSUBTYPE = "&sParSubType&"  "&_
							 "ORDER BY 1 "

		else
			'oDataList.sSQL = "select PartyName,PartyCode,OrgnPartyCode,PartySubType,PartyType from VwOrgParty "&_
			'				 "where OUDefinitionID='"&sOrgId&"' and PartyType='"&sParType&"' and PartySubType="&sParSubType&" "&_
			'				 "and PartyName Like '"&sFilter&"' order by PartyName"

			oDataList.sSQL = "SELECT V.PARTYNAME, "&_
							 "ISNULL(DBO.GETPARTYOPEN('"&sFinMonYear&"','"&sCurrDate&"','"&sOrgID&"','"&sParType&"',"&sParSubType&",V.PARTYCODE),'0.00') BALANCE, "&_
							 "V.PARTYCODE,V.ORGNPARTYCODE, V.PARTYSUBTYPE,V.PARTYTYPE FROM VWORGPARTY V WHERE  "&_
							 "V.OUDEFINITIONID = '"&sOrgID&"' AND PARTYTYPE = '"&sParType&"' AND PARTYSUBTYPE = "&sParSubType&"  "&_
							 "and PartyName Like '"&sFilter&"'  ORDER BY 1 "

		end if
	end if
Elseif instr(1,Cstr(sSearchBy),"IA")  Then
	IF CStr(sParSubType) = "0" Then
		if sFilter = "%" then
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"'and PartyType='"&sParType&"' order by PartyName"
		else
			sFilter = "%"&sFilter
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' and PartyType='"&sParType&"'  "&_
							 "and PartyName Like '"&sFilter&"' order by PartyName"
		End IF
	Elseif CStr(sParType) = "CRDR" and Cstr(sParSubType) = "000" Then
		if sFilter = "%" then
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' order by PartyName"
		else
			sFilter = "%"&sFilter
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"'   "&_
							 "and PartyName Like '"&sFilter&"' order by PartyName"
		End IF
	Else
		if sFilter = "%" then
			'oDataList.sSQL = "select PartyName,PartyCode,OrgnPartyCode,PartySubType,PartyType from VwOrgParty "&_
			'				 "where OUDefinitionID='"&sOrgId&"'and PartyType='"&sParType&"' and PartySubType="&sParSubType& " order by PartyName"

			oDataList.sSQL = "SELECT V.PARTYNAME, "&_
							 "ISNULL(DBO.GETPARTYOPEN('"&sFinMonYear&"','"&sCurrDate&"','"&sOrgID&"','"&sParType&"',"&sParSubType&",V.PARTYCODE),'0.00') BALANCE, "&_
							 "V.PARTYCODE,V.ORGNPARTYCODE, V.PARTYSUBTYPE,V.PARTYTYPE FROM VWORGPARTY V WHERE  "&_
							 "V.OUDEFINITIONID = '"&sOrgID&"' AND PARTYTYPE = '"&sParType&"' AND PARTYSUBTYPE = "&sParSubType&"  "&_
							 "ORDER BY 1 "
		else
			sFilter = "%"&sFilter
			'oDataList.sSQL = "select PartyName,PartyCode,OrgnPartyCode,PartySubType,PartyType from VwOrgParty "&_
			'				 "where OUDefinitionID='"&sOrgId&"' and PartyType='"&sParType&"' and PartySubType="&sParSubType&" "&_
			'				 "and PartyName Like '"&sFilter&"' order by PartyName"

			oDataList.sSQL = "SELECT V.PARTYNAME, "&_
							 "ISNULL(DBO.GETPARTYOPEN('"&sFinMonYear&"','"&sCurrDate&"','"&sOrgID&"','"&sParType&"',"&sParSubType&",V.PARTYCODE),'0.00') BALANCE, "&_
							 "V.PARTYCODE,V.ORGNPARTYCODE, V.PARTYSUBTYPE,V.PARTYTYPE FROM VWORGPARTY V WHERE  "&_
							 "V.OUDEFINITIONID = '"&sOrgID&"' AND PARTYTYPE = '"&sParType&"' AND PARTYSUBTYPE = "&sParSubType&"  "&_
							 "and PartyName Like '"&sFilter&"' ORDER BY 1 "

		end if
	end if
End IF

'Response.Write sQuery &"<br><br>"

'oDataList.sSQL = sQuery




'Response.Write oDataList.sSQL
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>
</BODY>
</HTML>