<%Option Explicit%>

<%
    'Program Name				:	GLHeadSelection.asp
	'Module Name				:	Common
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 24,2011
	'Modified On				:
%>
<!--#INCLUDE FILE="../include/clsDatabase.asp"-->
<!--#INCLUDE FILE="../include/clsDataListGeneric.asp"-->
<!--#include file="../include/populate.asp"-->
<HTML><HEAD><TITLE>GL Head Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sQuery
dim sIType,sOrgID,sFilter,sBookCode,iBookNo,iBkAccHead,sGroupCode,sSelectMode,sBalance
Dim sFinMonYear,sCurrDay,sCurrMon,sCurrDate
sOrgID = trim(Request.QueryString("orgID"))
sBookCode = trim(Request.QueryString("BookId"))
iBookNo = trim(Request.QueryString("BookNo"))
iBkAccHead = trim(Request.QueryString("AccHead"))
sFilter = trim(Request.QueryString("Query"))&"%"
sGroupCode = trim(Request.QueryString("GroupCode"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
sBalance = Trim(Request.QueryString("hBal"))
if trim(sGroupCode)<>"" then sGroupCode = "'"& sGroupCode &"'"
if trim(sBalance)="" then sBalance = "N"


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
if trim(sBalance)<>"N" then
    oDataList.AddDisplayField "Account Head Description"
    oDataList.AddDisplayField "Balance"
else
    oDataList.AddDisplayField "Account Head Description"
end if

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
    if trim(sGroupCode)<>"" then
        oDataList.sSQL = " Select V.AccountDescription,isNull(dbo.GetGLOpen('"&sFinMonYear&"','"&sCurrDate&"','0','"&sOrgID&"',V.AccountHead),'0.00') Balance, "&_
				         "V.AccountHead,V.AccountHeadCode,isnull(V.CostCenterExists,'0'),isnull(V.AnalyticalHeadExists,'0'),V.AllowTransactions,V.EligibleForTds "&_
				         "From VwOrgGLHeads V Where V.OUDefinitionID = '"&sOrgID&"' and V.AccountGroupCode in("& sGroupCode &")  And V.SubLedger = '0' Order BY 1  "
    else
	    oDataList.sSQL = " Select V.AccountDescription,isNull(dbo.GetGLOpen('"&sFinMonYear&"','"&sCurrDate&"','0','"&sOrgID&"',V.AccountHead),'0.00') Balance, "&_
				         "V.AccountHead,V.AccountHeadCode,isnull(V.CostCenterExists,'0'),isnull(V.AnalyticalHeadExists,'0'),V.AllowTransactions,V.EligibleForTds "&_
				         "From VwOrgGLHeads V Where V.OUDefinitionID = '"&sOrgID&" ' And V.SubLedger = '0' Order BY 1  "
    end if 
else
    if trim(sGroupCode)<>"" then
        oDataList.sSQL = " Select V.AccountDescription,isNull(dbo.GetGLOpen('"&sFinMonYear&"','"&sCurrDate&"','0','"&sOrgID&"',V.AccountHead),'0.00') Balance, "&_
	    			     "V.AccountHead,V.AccountHeadCode,isnull(V.CostCenterExists,'0'),isnull(V.AnalyticalHeadExists,'0'),V.AllowTransactions,V.EligibleForTds "&_
		    		     "From VwOrgGLHeads V Where V.OUDefinitionID = '"&sOrgID&" ' And V.SubLedger = '0' and V.AccountGroupCode in("& sGroupCode &") and AccountDescription Like '"&sFilter&"'  Order BY 1  "
    else
    	oDataList.sSQL = " Select V.AccountDescription,isNull(dbo.GetGLOpen('"&sFinMonYear&"','"&sCurrDate&"','0','"&sOrgID&"',V.AccountHead),'0.00') Balance, "&_
			    	     "V.AccountHead,V.AccountHeadCode,isnull(V.CostCenterExists,'0'),isnull(V.AnalyticalHeadExists,'0'),V.AllowTransactions,V.EligibleForTds "&_
				         "From VwOrgGLHeads V Where V.OUDefinitionID = '"&sOrgID&" ' And V.SubLedger = '0' and AccountDescription Like '"&sFilter&"'  Order BY 1  "
    end if 
end if
'Response.Write oDataList.sSQL
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>
</BODY>
</HTML>