
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PartySelMultipleSubType.asp
	'Module Name				:	ACCOUNTS
	'Author Name				:	S.Maheswari
	'Created On					:	FEB 26, 2009
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
<!--#include virtual="/include/clsDatalistParSubTypeMultiple.asp"-->
<!--#include virtual="/include/populate.asp"-->
<HTML><HEAD><TITLE>Party Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sQuery,sTemp
dim sIType,sOrgID,sFilter,sParType,sParSubType,sParTypeName,sSearchBy,sFinMonYear,sCurrDate
Dim sCurrDay,sCurrMon,sSelectMode,sParVal
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
 
'IF trim(Request("Party")) <> "" then 
sParVal  = trim(Request("Party"))
'sTemp=split(trim(Request("Party")),"?")
'Response.write sParVal
 dim sArr,i,sNewParType
IF trim(sParVal) <> "0" then
	sTemp=split(sParVal,"?")
	sNewParType=sTemp(0)
	IF len(sNewParType) > 2 then
		sArr = split(sNewParType,",") 
		for i = 0 to uBound(sArr)
			sParType = sParType &",'"&  sArr(i)&"'"
		Next 
		
	End IF
	sParType=mid(sParType,2)  
	sParSubType=sTemp(1)
	sParTypeName=sTemp(2)
	
Else
	sParType=""
	sParSubType="0"
	sParTypeName=""
End IF

sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode"))) 
 
'Response.Write sParType&" " & sParSubType &"<br>"

sFilter = trim(Request.QueryString("Query"))&"%"

'Response.Write "sSearchBy = "& sSearchBy &"<br>"
'Response.Write "sFilter = "& sFilter &"<br>"


Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist

oDataList.PageSize = 15

oDataList.PrimaryKey = "PartyCode"
oDataList.AddOptDispField sSelectMode
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


	oDataList.AddReturnedField "0"
	oDataList.AddReturnedField "1"
	oDataList.AddReturnedField "2"
	oDataList.AddReturnedField "3"
	oDataList.AddReturnedField "4"


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
							 "where OUDefinitionID='"&sOrgId&"' "
			IF  trim(sParType) <> "" then
				oDataList.sSQL = oDataList.sSQL & " and PartyType in ("&sParType&") " 
			End IF
			
		else
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' and PartyName Like '"&sFilter&"' "
			IF  trim(sParType) <> "" then
				oDataList.sSQL = oDataList.sSQL & " and PartyType in ("&sParType&")  "
			End IF
		End IF
		oDataList.sSQL = oDataList.sSQL & "order by PartyName"
	'Elseif CStr(sParType) = "CRDR" and Cstr(sParSubType) = "000" Then
	Elseif CStr(sParType) = "" and Cstr(sParSubType) = "0" Then
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
		
			oDataList.sSQL = "SELECT PARTYNAME,PARTYCODE,ORGNPARTYCODE, PARTYSUBTYPE,PARTYTYPE FROM VWORGPARTY  WHERE  "&_
							 "OUDEFINITIONID = '"&sOrgID&"' "
				IF  trim(sParType) <> "" then				 
					oDataList.sSQL = oDataList.sSQL & " AND PARTYTYPE in("&sParType&" ) "
				End If
				IF  trim(sParSubType) <> "0" then
					oDataList.sSQL = oDataList.sSQL & " AND PARTYSUBTYPE in( "&sParSubType&" )"
				End IF				  
		else
			
			oDataList.sSQL = "SELECT PARTYNAME,PARTYCODE,ORGNPARTYCODE, PARTYSUBTYPE,PARTYTYPE FROM VWORGPARTY  WHERE  "&_
							 "OUDEFINITIONID = '"&sOrgID&"' and PartyName Like '"&sFilter&"'  "
				IF  trim(sParType) <> "" then				 
					oDataList.sSQL = oDataList.sSQL & " AND PARTYTYPE in("&sParType&") "
				End If
				IF  trim(sParSubType) <> "0" then
					oDataList.sSQL = oDataList.sSQL & " AND PARTYSUBTYPE in( "&sParSubType&" ) "
				End IF
		end if
		oDataList.sSQL = oDataList.sSQL & " ORDER BY 1 "
	end if
Elseif instr(1,Cstr(sSearchBy),"IA")  Then
	IF CStr(sParSubType) = "0" Then
		if sFilter = "%" then
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' "
				IF  trim(sParType) <> "" then				 
					oDataList.sSQL = oDataList.sSQL & "and PartyType in ("&sParType&")"
				End IF
		else
			sFilter = "%"&sFilter
			oDataList.sSQL = "select Distinct PartyName,PartyCode,OrgnPartyCode,OrgnPartyCode,OrgnPartyCode from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' and PartyName Like '"&sFilter&"' "
				IF  trim(sParType) <> "" then				 
					oDataList.sSQL = oDataList.sSQL & " and PartyType in ("&sParType&") "
				End If				 
							 
		End IF
		oDataList.sSQL = oDataList.sSQL & " order by PartyName "
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
			
			oDataList.sSQL = "SELECT PARTYNAME,PARTYCODE,ORGNPARTYCODE, PARTYSUBTYPE,PARTYTYPE FROM VWORGPARTY  WHERE  "&_
							 "OUDEFINITIONID = '"&sOrgID&"' "
				IF trim(sParType) <> "" then 
					oDataList.sSQL = oDataList.sSQL & " AND PARTYTYPE in("&sParType&") " 
				End IF
				IF trim(PARTYSUBTYPE) <> "" then 
					oDataList.sSQL = oDataList.sSQL & " AND PARTYSUBTYPE in("&sParSubType&")"
				End IF
							 
		else
			sFilter = "%"&sFilter
			
			oDataList.sSQL = "SELECT PARTYNAME,PARTYCODE,ORGNPARTYCODE,PARTYSUBTYPE,PARTYTYPE FROM VWORGPARTY  WHERE  "&_
							 "OUDEFINITIONID = '"&sOrgID&"' and PartyName Like '"&sFilter&"'  "
				IF trim(sParType) <> "" then 
					oDataList.sSQL = oDataList.sSQL & " AND PARTYTYPE in("&sParType&") "
				End IF
				IF trim(PARTYSUBTYPE) <> "" then 
					oDataList.sSQL = oDataList.sSQL & " AND PARTYSUBTYPE in("&sParSubType&") "
				End IF
		end if
		oDataList.sSQL = oDataList.sSQL & " ORDER BY 1 "	
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