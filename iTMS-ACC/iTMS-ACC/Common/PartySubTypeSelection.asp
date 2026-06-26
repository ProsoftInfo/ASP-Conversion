<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PartySubTypeSelection.asp
	'Module Name				:	Common
	'Author Name				:	UmaMaheswari S
	'Created On					:	August 05,2011
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
<!--#INCLUDE FILE="../include/clsDatabase.asp"-->
<!--#INCLUDE FILE="../include/clsDataListGeneric.asp"-->
<!--#include file="../include/populate.asp"-->
<HTML><HEAD><TITLE>Party SubType Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sQuery,sTemp
dim sIType,sOrgID,sFilter,sParType,sParSubType,sParTypeName,sSearchBy,sFinMonYear,sCurrDate
Dim sCurrDay,sCurrMon,sSelectMode,sTempData
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
'sTemp = trim(Request("Party"))
sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
Response.Write "<p><font color=red>"
'Response.Write "Data="&trim(Request("Party"))


if trim(Request("Party"))<>"" then
    if UBound(sTemp)>1 then
	    sParType=sTemp(0)
        sParSubType=sTemp(1)
        sParTypeName=sTemp(2)
    else
        sParType = sTemp(0)
    end if
end if 'if trim(Request("Party"))<>"" then

sFilter = trim(Request.QueryString("Query"))&"%"

Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist

oDataList.PageSize = 15

oDataList.PrimaryKey = "PartySubType"
oDataList.AddOptDispField sSelectMode

oDataList.AddDisplayField "Party SubType Name"

oDataList.AddReturnedField "0"
oDataList.AddReturnedField "1"
oDataList.AddReturnedField "2"
oDataList.AddReturnedField "3"
oDataList.AddReturnedField "4"
oDataList.AddReturnedField "5"	'Added By UmaMaheswari On 5th MAY 2011,For CtrlAccCheck
	

oDataList.SearchForDesc "Search For"
oDataList.AddSearchField "Party Sub Name < Starts with >","IN"
oDataList.AddSearchField "Party Sub Name < Anywhere >","IA"

IF CStr(sSearchBy) = "" Then
	sSearchBy = "IN"
End IF
'Response.Write "sParType = "& sParType 

IF instr(1,Cstr(sSearchBy),"IN") > 0  Then
		if sFilter = "%" then
			oDataList.sSQL = "select Distinct SubTypeName,PartySubType,0,0,0,PartySubType from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' order by SubTypeName"
							 
		    if trim(sParType)<>"" then
		        oDataList.sSQL = "select Distinct SubTypeName,PartySubType,0,0,0,PartySubType from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"'and PartyType='"&sParType&"' order by SubTypeName"
		    end if
		    
		else
			oDataList.sSQL = "select Distinct SubTypeName,PartySubType,0,0,0,PartySubType from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"'"&_
							 "and SubTypeName Like '"&sFilter&"' order by SubTypeName"
							 
		    if trim(sParType)<>"" then
		        oDataList.sSQL = "select Distinct SubTypeName,PartySubType,0,0,0,PartySubType from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' and PartyType='"&sParType&"' and SubTypeName Like '"&sFilter&"' order by SubTypeName"
		    end if
		    
		End IF
Elseif instr(1,Cstr(sSearchBy),"IA")  Then
		if sFilter = "%" then
			oDataList.sSQL = "select Distinct SubTypeName,PartySubType,0,0,0,PartySubType from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"'  order by SubTypeName"
			if trim(sParType)<>"" then
			    oDataList.sSQL = "select Distinct SubTypeName,PartySubType,0,0,0,PartySubType from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' and PartyType='"&sParType&"' order by SubTypeName"
			end if
			
		Else
			sFilter = "%"&sFilter
			oDataList.sSQL = "select Distinct SubTypeName,PartySubType,0,0,0,PartySubType from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"'"&_
							 "and SubTypeName Like '"&sFilter&"' order by SubTypeName"
			if trim(sParType)<>"" then
			    oDataList.sSQL = "select Distinct SubTypeName,PartySubType,0,0,0,PartySubType from VwOrgParty "&_
							 "where OUDefinitionID='"&sOrgId&"' and PartyType='"&sParType&"' and SubTypeName Like '"&sFilter&"' order by SubTypeName"
			end if
			
		End IF
End IF

'oDataList.sSQL = sQuery

'Response.Write "<p><font color=red>"&oDataList.sSQL 
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>
</BODY>
</HTML>