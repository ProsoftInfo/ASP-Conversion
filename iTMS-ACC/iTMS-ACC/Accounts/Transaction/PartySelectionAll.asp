<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PartySelectionAll.asp
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
dim sIType,sOrgID,sFilter,sParType,sParSubType,sParTypeName,sSearchBy

sOrgID = trim(Request.QueryString("orgID"))
sTemp=split(trim(Request("Party")),"?")
sSearchBy = Trim(Request("SearchBy"))

IF UBound(sTemp) > 1 Then
	sParType=sTemp(0)
	sParSubType=sTemp(1)
	sParTypeName=sTemp(2)
End IF

sFilter = trim(Request.QueryString("Query"))&"%"

'Response.Write sFilter &"<br>"


Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist

oDataList.PageSize = 15

oDataList.PrimaryKey = "PartyCode"



oDataList.AddDisplayField "Party Name"

oDataList.AddReturnedField "0"
oDataList.AddReturnedField "1"
oDataList.AddReturnedField "2"
oDataList.AddReturnedField "1"
oDataList.AddReturnedField "1"

oDataList.SearchForDesc "Search For"
oDataList.AddSearchField "Party Name < Starts with >","IN"
oDataList.AddSearchField "Party Name < Anywhere >","IA"

IF CStr(sSearchBy) <> "" Then
	IF Cstr(sSearchBy) = "IN"  Then
		sFilter = sFilter&"%"
		'oDataList.sSQL = "Select PartyName,PartyCode,OrgnPartyCode From App_M_PartyMaster Where Useable = '0' and PartyName Like " & Pack(sIType) & " Order By PartyName "
		
		sQuery = "Select PartyName,PartyCode,OrgnPartyCode From App_M_PartyMaster Where Useable = '0' and PartyName Like " & Pack(sFilter) & " Order By PartyName "
	elseif Cstr(sSearchBy) = "IA"  Then
		sFilter = "%"&sFilter&"%"
		'oDataList.sSQL = "Select PartyName,PartyCode,OrgnPartyCode From App_M_PartyMaster "&_
		'		 "Where PartyName Like " & Pack(sIType) & " and Useable = '0' order by PartyName"
				 
		sQuery = "Select PartyName,PartyCode,OrgnPartyCode From App_M_PartyMaster "&_
				 "Where PartyName Like " & Pack(sFilter) & " and Useable = '0' order by PartyName"
		
	Else
		'oDataList.sSQL = "Select PartyName,PartyCode,OrgnPartyCode From App_M_PartyMaster Where Useable = '0' Order By PartyName "
		sQuery = "Select PartyName,PartyCode,OrgnPartyCode From App_M_PartyMaster Where Useable = '0' Order By PartyName "
	End IF
Else
	'oDataList.sSQL = "Select PartyName,PartyCode,OrgnPartyCode From App_M_PartyMaster Where Useable = '0' Order By PartyName "
	sQuery = "Select PartyName,PartyCode,OrgnPartyCode From App_M_PartyMaster Where Useable = '0' Order By PartyName "
End IF

'Response.Write sQuery &"<br><br>"

oDataList.sSQL = sQuery

sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>
</BODY>
</HTML>