<%option Explicit%>
<!--#INCLUDE FILE="../../include/clsSuppDataListMultiple.asp"-->
<!--#INCLUDE FILE="../../include/clsDatabase.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/purpopulate.asp"-->

<HTML><HEAD><TITLE>Party Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%
	Dim oDatabase,oDataList,rsTemp
	
	Dim sTable,sFilter,sSelectMode,sSqlQuery,sSearchBy,sUnit,sItemType
	
	Dim nFlag
	'Dim nAgentCode
	
	set rsTemp = server.CreateObject("ADODB.RecordSet")
	
	'Response.Write "<p> para = "  & Request.QueryString
	sSearchBy =trim(Request.QueryString("SearchBy"))
	sFilter = trim(Request.QueryString("Query"))	
	sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
	nFlag = trim(Request.QueryString("Flag"))
	
	sUnit		= trim(Request.QueryString("Unit")) 
	sItemType	= trim(Request.QueryString("ItemType"))
	
	'Response.Write "<p> "  & Request.QueryString 	
	
	if trim(sSelectMode) = "" then sSelectMode = "R"
	
	If trim(sFilter) <> "" Then
		sFilter = sFilter  &"%"
	End If
	
	
	' note : based on nflag value Sql Query will form
	If nFlag = "" Then nFlag = 0
	' note: about nflag Value & corresoponding reference details
	'		 nflag	= 1 -- Gate Receipt entry
	'				= 2 -- Receipt Status - Data list
	
	'Response.Write "<p> nFlag =  "& nFlag
	Set oDatabase = New clsDatabase
	Set oDatalist = New clsDatalist
		
	oDataList.PageSize = 15
	oDataList.PrimaryKey = "PartyCode"
	oDataList.AddDisplayField "Party Name"
	oDataList.AddDisplayField "Party Sub.Type Name"
	
	oDataList.AddOptDispField sSelectMode
		
	oDataList.AddReturnedField "0"
	oDataList.AddReturnedField "1"
	oDataList.AddReturnedField "2"
	oDataList.AddReturnedField "3"
	oDataList.AddReturnedField "3"
	oDataList.AddReturnedField "4"
	
	oDataList.SearchForDesc "Search For"
	oDataList.AddSearchField "Party Name < Starts with >","SN"
	oDataList.AddSearchField "Party Sub.Type Name","ST"
	oDataList.AddSearchField "Party Name < Anywhere >","SA"
	 
	
	'Response.Write "<br>" & sSearchBy
	 		
			
	if sFilter = "" then
		if nFlag = 1 then
			sSqlQuery =     "Select distinct PartyName,SubTypeName,PartyCode,PartyType,PartySubType from VwOrgParty where PartyCode <> 0 order by PartyName"
		elseif nFlag = 2 then
		
			'Select Distinct A.PartyName,E.SubTypeName,A.PartyCode,B.PartyType,B.PartySubType from App_M_PartyMaster A,
			'APP_R_OrgParty B , Rcv_T_GateReceiptHeader C,RCV_T_GRNItemDetails D, APP_M_PartyTypes E Where B.PartyType = E.PartyType and B.PartySubType = E.PartySubType  and b.PartyCode=a.PartyCode and D.GRNNumber=C.GRNNumber and D.ItemTypeID='"&sItemType&"' and 
			'A.PartyCode = C.PartyCode 
			'and B.PartyType = C.PartyType and B.PartySubType = C.ReceivedFrom
			'and C.ReceivedForUnit='" & trim(sUnit) & "' and B.OUDefinitionID=C.ReceivedForUnit order by A.PartyName 
			
			sSqlQuery = "Select Distinct A.PartyName,E.SubTypeName,A.PartyCode,B.PartyType,B.PartySubType from App_M_PartyMaster A, " & _
						"APP_R_OrgParty B , Rcv_T_GateReceiptHeader C,RCV_T_GRNItemDetails D, APP_M_PartyTypes E Where B.PartyType = E.PartyType and B.PartySubType = E.PartySubType  and b.PartyCode=a.PartyCode and D.GRNNumber=C.GRNNumber "
			if trim(sItemType) <> "" then
				sSqlQuery = sSqlQuery & " and D.ItemTypeID='"&sItemType&"'"
			end if 			
			sSqlQuery = sSqlQuery & " and A.PartyCode = C.PartyCode " &_
						" and C.ReceivedForUnit='" & trim(sUnit) & "' and B.OUDefinitionID=C.ReceivedForUnit order by A.PartyName"
		end if 'if nFlag = 1 then
	elseif sFilter <> "" then
		if nFlag = 1 then
			sSqlQuery = "Select distinct PartyName,SubTypeName,PartyCode,PartyType,PartySubType from VwOrgParty"
			if trim(sSearchBy) = "ST" then
				sFilter = "%" & sFilter & "%"
				sSqlQuery = sSqlQuery & " where SubTypeName Like " & Pack(trim(sFilter)) & " order by 2"
			elseif trim(sSearchBy) = "SN" then
				sFilter = sFilter & "%"
				sSqlQuery = sSqlQuery & " where PartyName like " & Pack(sFilter) & ""
			elseif trim(sSearchBy) = "SA" then
				sFilter = "%" & sFilter & "%"
				sSqlQuery = sSqlQuery & " where PartyName Like " & Pack(sFilter) & ""
			end if
		elseif nFlag = 2 then
		
			sSqlQuery = "Select Distinct A.PartyName,E.SubTypeName,A.PartyCode,B.PartyType,B.PartySubType from App_M_PartyMaster A, " & _
						"APP_R_OrgParty B , Rcv_T_GateReceiptHeader C,RCV_T_GRNItemDetails D, APP_M_PartyTypes E Where B.PartyType = E.PartyType and B.PartySubType = E.PartySubType  and b.PartyCode=a.PartyCode and D.GRNNumber=C.GRNNumber "
			if trim(sItemType) <> "" then
				sSqlQuery = sSqlQuery & " and D.ItemTypeID='"&sItemType&"'"
			end if 			
			sSqlQuery = sSqlQuery & " and A.PartyCode = C.PartyCode " &_
						" and C.ReceivedForUnit='" & trim(sUnit) & "' and B.OUDefinitionID=C.ReceivedForUnit "
						
			if trim(sSearchBy) = "ST" then
				sFilter = "%" & sFilter & "%"
				sSqlQuery = sSqlQuery & " and E.SubTypeName Like " & Pack(trim(sFilter)) & " order by 2"
			elseif trim(sSearchBy) = "SN" then
				sFilter = sFilter & "%"
				sSqlQuery = sSqlQuery & " and A.PartyName like " & Pack(sFilter) & ""
			elseif trim(sSearchBy) = "SA" then
				sFilter = "%" & sFilter & "%"
				sSqlQuery = sSqlQuery & " and A.PartyName Like " & Pack(sFilter) & ""
			end if
		end if 'if nFlag = 1 then	
	end if	'if sFilter = "" then
		
		
	sFilter = "%" & trim(sFilter) & "%"
		
	oDataList.sSQL = sSqlQuery
	
	'Response.Write "<p> sSqlQuery = " & sSqlQuery
	
	sTable = oDatalist.GetTable(oDatabase)
	
	Set oDatalist = Nothing
	Set oDatabase = Nothing
	Response.Write sTable
%>
</BODY>
</HTML>
