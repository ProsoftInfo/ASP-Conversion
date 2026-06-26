<%@ Language=VBScript %>
<%option Explicit%>
<%
	'Program Name				:	ParentPartySelection.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 19,2010
	'Modified By				:	
	'Modified By				:
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

<!--#INCLUDE FILE="../../include/clsSuppDataListMultiple.asp"-->
<!--#INCLUDE FILE="../../include/clsDatabase.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/purpopulate.asp"-->

<HTML><HEAD><TITLE>Supplier Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%
	Dim oDatabase,oDataList,sTable,sFilter,sSelectMode
	Dim nFlag,rsTemp,sSearchBy,sOrgID,sOrgIDVal,sSql
	Dim sRef,sOrderTo
	
	'Dim nAgentCode
	
	set rsTemp = server.CreateObject("ADODB.RecordSet")
	
	'Response.Write "<p> para = "  & Request.QueryString
	sSearchBy =trim(Request.QueryString("SearchBy"))
	sFilter = trim(Request.QueryString("Query"))	
	sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
	
'	Response.Write "sFilter="& sFilter
	
	if trim(sSelectMode) = "" then sSelectMode = "R"
	
	If trim(sFilter) <> "" Then
		sFilter = sFilter  &"%"
	End If
	
		
	'Response.Write "<p> nFlag =  "& nFlag
	Set oDatabase = New clsDatabase
	Set oDatalist = New clsDatalist
		
	oDataList.PageSize = 15
	oDataList.PrimaryKey = "partycode"
	oDataList.AddDisplayField "Party Code"
	oDataList.AddDisplayField "Party Name"
	
	oDataList.AddOptDispField sSelectMode
		
	oDataList.AddReturnedField "0" 'OrgnPartyCode
	oDataList.AddReturnedField "1" 'PartyCode
	oDataList.AddReturnedField "2" 'PartyName
	oDataList.AddReturnedField "3" 
	oDataList.AddReturnedField "4" 
	oDataList.AddReturnedField "5"
	
	oDataList.SearchForDesc "Search For"
	oDataList.AddSearchField "Party Name < Starts with >","SN"
	oDataList.AddSearchField "Party Name < Anywhere >","SA"
	 
		if trim(sFilter) = "" then
			oDataList.sSQL = "Select OrgnPartyCode,PartyName,PartyCode,0,0,0 from APP_M_PartyMaster where PartyGroupCoyType='P' or PartyGroupCoyType='B'"
		elseif sFilter <> "" then
			if sSearchBy ="SN" then
				oDataList.sSQL = "Select OrgnPartyCode,PartyName,PartyCode,0,0,0 from APP_M_PartyMaster where (PartyGroupCoyType='P' or PartyGroupCoyType='B')  and PartyName = '" & sFilter&"'"
			elseif sSearchBy = "SA" then
				oDataList.sSQL = "Select OrgnPartyCode,PartyName,PartyCode,0,0,0 from APP_M_PartyMaster where (PartyGroupCoyType='P' or PartyGroupCoyType='B')  and PartyName = '%" & sFilter&"'"
			end if
		end if	
		
'		Response.Write oDataList.sSQL
		
	sTable = oDatalist.GetTable(oDatabase)
	Set oDatalist = Nothing
	Set oDatabase = Nothing
	Response.Write sTable
%>
</BODY>
</HTML>


