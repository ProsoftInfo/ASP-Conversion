<%Option Explicit%>

<%
    'Program Name				:	BookSelectionPopUp.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	UmaMaheswari S
	'Created On					:	Mar, 26 2011
	'Modified On				:
%>
<!--#include virtual="/include/clsDatabase.asp"-->
<!--#include virtual="/include/clsDataListGeneric.asp"-->
<!--#include virtual="/include/populate.asp"-->
<HTML><HEAD><TITLE>Book Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sQuery
dim sIType,sOrgID,sFilter,sBookCode,iBookNo,iBkAccHead,sGroupCode,sSelectMode

sOrgID = trim(Session("organizationcode"))

sFilter = trim(Request.QueryString("Query"))&"%"
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
sBookCode = trim(Request.QueryString("BookCode"))

Set oDatabase = New clsDatabase
Set oDatalist = New clsDataList

oDataList.PageSize = 15

oDataList.PrimaryKey = "BookCode"
oDataList.AddDisplayField "Book Name"
oDataList.AddOptDispField sSelectMode
oDataList.AddReturnedField "2"
oDataList.AddReturnedField "3"
oDataList.AddReturnedField "0"
oDataList.AddReturnedField "0"
oDataList.AddReturnedField "0"

oDataList.SearchForDesc "Search For"
oDataList.AddSearchField "Book Name","AD"
oDataList.AddOptDispField sSelectMode

if sFilter = "%" then
    oDataList.sSQL = " select BookName,OUDefinitionID,BookCode,BookNumber,OrgUnitShortDescription,BookName,0,0,0,0 "&_
                     " from vwOrgBookNames Where BookCode="& sBookCode &" Order By OrgUnitShortDescription,BookName "
else
	oDataList.sSQL = " select BookName,OUDefinitionID,BookCode,BookNumber,OrgUnitShortDescription,BookName,0,0,0,0 "&_
                     " from vwOrgBookNames where BookCode="& sBookCode &" and BookName Like '"&sFilter&"' Order By BookName "
end if
'Response.Write oDataList.sSQL
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>
</BODY>
</HTML>