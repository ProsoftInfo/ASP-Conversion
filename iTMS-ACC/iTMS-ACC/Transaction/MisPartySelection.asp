<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MisPartySelection.asp
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
<!--#INCLUDE FILE="../../include/clsDatabase.asp"-->
<!--#INCLUDE FILE="../../include/clsMisParDatalist.asp"-->
<!--#include file="../../include/populate.asp"-->
<HTML><HEAD><TITLE>Party Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sQuery,sTemp
dim sIType,sOrgID,sFilter,sParType,sParSubType,sParTypeName

sOrgID = trim(Request.QueryString("orgID"))
'sTemp=split(trim(Request("Party")),"?")

'sParType=sTemp(0)
'sParSubType=sTemp(1)
'sParTypeName=sTemp(2)

sFilter = trim(Request.QueryString("Query"))&"%"




Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist

oDataList.PageSize = 15

oDataList.PrimaryKey = "MiscPartyCode"


oDataList.AddDisplayField "Party Name"

oDataList.AddReturnedField "0"
oDataList.AddReturnedField "1"
oDataList.AddReturnedField "1"
oDataList.AddReturnedField "1"
oDataList.AddReturnedField "1"


if sFilter = "%" then
	oDataList.sSQL = "SELECT PartyName,MiscPartyCode  FROM App_M_MiscPartyMaster Order By PartyName "
else
	oDataList.sSQL = "SELECT PartyName,MiscPartyCode  FROM App_M_MiscPartyMaster Where "&_
					 "PartyName Like '"&sFilter&"' order by PartyName"
End IF




sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>
</BODY>
</HTML>