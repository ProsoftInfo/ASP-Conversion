<%Option Explicit%>

<%
    'Program Name				:	NarrSelection.asp
	'Module Name				:	Common
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 27,2011
	'Modified On				:
%>
<!--#INCLUDE FILE="../include/clsDatabase.asp"-->
<!--#INCLUDE FILE="../include/clsDataListGeneric.asp"-->
<!--#include file="../include/populate.asp"-->
<HTML><HEAD><TITLE>GL Head Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sQuery
dim sIType,sOrgID,sFilter,sBookCode,iBookNo,iBkAccHead,sGroupCode,sSelectMode,sBalance
Dim sFinMonYear,sCurrDay,sCurrMon,sCurrDate
sOrgID = trim(Request.QueryString("orgID"))

sBookCode = trim(Request.QueryString("BookId"))
iBookNo = trim(Request.QueryString("BookNo"))
sFilter = trim(Request.QueryString("Query"))&"%"

sFinMonYear = "04"
sFinMonYear = sFinMonYear&Trim(Left(Session("FinPeriod"),4))
sCurrDay = Day(Now)
sCurrMon = Month(Now)
'Response.Write sSelectMode
'Response.End 

IF Len(Trim(sCurrDay)) = 1 Then
	sCurrDay = "0"&Trim(sCurrDay)
End IF

IF Len(Trim(sCurrMon)) = 1 Then
	sCurrMon = "0"&Trim(sCurrMon)
End IF

sCurrDate = Trim(sCurrDay&"/"&sCurrMon&"/"&Year(Now))


Set oDatabase = New clsDatabase
Set oDatalist = New clsDataList

oDataList.PageSize = 15

oDataList.PrimaryKey = "AccountHead"

'oDataList.AddDisplayField "Item Code"
    oDataList.AddDisplayField "Short Narr."
    oDataList.AddDisplayField "Narration"

    oDataList.AddReturnedField "2" ' Account Head  ' RetField0
    oDataList.AddReturnedField "3" ' Account Head Code ' RetField1
    oDataList.AddReturnedField "4" ' Cost Center Exists ' RetField2
    oDataList.AddReturnedField "5" ' Analytical Head Exists ' RetField3
    oDataList.AddReturnedField "6" ' Allow Transaction ' RetField4
    oDataList.AddReturnedField "0" ' Account Description ' RetField5
    oDataList.AddReturnedField "7" ' Eligible For TDS ' RetField6
    oDataList.AddReturnedField "1" ' Balance ' RetField7
    

oDataList.SearchForDesc "Search For"
oDataList.AddSearchField "Account Description","AD"
oDataList.AddOptDispField sSelectMode

if sFilter = "%" then
        sQuery ="select NarrationDesc,NarrationNumber,NarrationShortDesc from VwOrgFrequentNarration where "&_
        	" OUDefinitionID='"&sOrgId&"'and BookCode='"&sBookCode&"' and BookNumber="&sBookNo
else
        sQuery ="select NarrationDesc,NarrationNumber,NarrationShortDesc from VwOrgFrequentNarration where "&_
        	" OUDefinitionID='"&sOrgId&"'and BookCode='"&sBookCode&"' and BookNumber="& sBookNo &" and NarrationDesc like '"& sFilter&"'"
end if
'Response.Write oDataList.sSQL
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>
</BODY>
</HTML>