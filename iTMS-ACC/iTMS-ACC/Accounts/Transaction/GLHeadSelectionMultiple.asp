<%Option Explicit%>
<!--#INCLUDE FILE="../../include/clsDatabase.asp"-->
<!--#INCLUDE FILE="../../include/clsGLDatalistMultiple.asp"-->
<!--#include file="../../include/populate.asp"-->
<HTML><HEAD><TITLE>GL Head Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sQuery
dim sIType,sOrgID,sFilter,sBookCode,iBookNo,iBkAccHead,sFinMonYear,sCurrDate
Dim sCurrDay,sCurrMon,sSelectMode,sSearchBy

sOrgID = trim(Request.QueryString("orgID"))
'sOrgID = trim(Request.QueryString("orgid"))
sOrgId = Session("organizationcode")
sBookCode = trim(Request.QueryString("BookId"))
iBookNo = trim(Request.QueryString("BookNo"))
iBkAccHead = trim(Request.QueryString("AccHead"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
sFilter = trim(Request.QueryString("Query"))
'&"%"
sSearchBy =   trim(Request.QueryString("SearchBy"))
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
'Response.Write sCurrDate 
'Response.Write sFinMonYear  

Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist

oDataList.PageSize = 14

oDataList.PrimaryKey = "AccountHead"
oDataList.AddOptDispField sSelectMode
'oDataList.AddDisplayField "Item Code"
oDataList.AddDisplayField "Account Head Description"
oDataList.AddDisplayField "  Balance  "

oDataList.AddReturnedField "2"
oDataList.AddReturnedField "3"
oDataList.AddReturnedField "4"
oDataList.AddReturnedField "5"
oDataList.AddReturnedField "6"
oDataList.AddReturnedField "0"
oDataList.AddReturnedField "7"
oDataList.AddReturnedField "1"
oDataList.SearchForDesc "Search For" 
oDataList.AddSearchField "AccHead Name < Starts with >","IN"
oDataList.AddSearchField "AccHead Name < Anywhere >","IA"
'Response.Write sFilter

if sFilter = "" then
	IF CStr(iBkAccHead) = "" Then
		'oDataList.sSQL = "select AccountDescription,AccountHead,AccountHeadCode,isnull(CostCenterExists,'0'),"&_
		'				 "isnull(AnalyticalHeadExists,'0'),AllowTransactions,EligibleForTds from VwOrgGLHeads where SubLedger=0 and OUDefinitionID='"&sOrgId&"'" & " order by AccountDescription "

		'sQuery = "SELECT V.AccountDescription, isNull(F.CloseAmtWithIndi,0.00),V.AccountHead, V.AccountHeadCode, "&_
		'		 "isNull(V.CostCenterExists,0),isNull(V.AnalyticalHeadExists,0), V.AllowTransactions,  "&_
		'		 "V.EligibleForTDS, V.OUDefinitionID  FROM dbo.VwOrgGLHeads V INNER JOIN "&_
		'		 "dbo.GetGLOpen1('"&sFinMonYear&"', '"&sCurrDate&"', '0', '"&sOrgID&"') F ON V.AccountHead = F.AccHD "&_
		'		 "WHERE V.SubLedger = '0' and V.OUDefinitionID = '"&sOrgID&"' ORDER BY 1"

		sQuery = "Select V.AccountDescription,isNull(dbo.GetGLOpen('"&sFinMonYear&"','"&sCurrDate&"','0','"&sOrgID&"',V.AccountHead),'0.00') Balance, "&_
				 "V.AccountHead,V.AccountHeadCode,isnull(V.CostCenterExists,'0'), "&_
				 "isnull(V.AnalyticalHeadExists,'0'),V.AllowTransactions,V.EligibleForTds "&_
				 "From VwOrgGLHeads V Where V.OUDefinitionID = '"&sOrgID&" ' And V.SubLedger = '0' Order BY 1  "
		
	Else
		'oDataList.sSQL = "select AccountDescription,AccountHead,AccountHeadCode,isnull(CostCenterExists,'0'),"&_
		'				 "isnull(AnalyticalHeadExists,'0'),AllowTransactions,EligibleForTds from VwOrgGLHeads where SubLedger=0 and OUDefinitionID='"&sOrgId&"' and AccountHead <> '"&iBkAccHead&"' "& " order by AccountDescription "

		'sQuery = "SELECT V.AccountDescription, isNull(F.CloseAmtWithIndi,0.00),V.AccountHead, V.AccountHeadCode, "&_
		'		 "isNull(V.CostCenterExists,0),isNull(V.AnalyticalHeadExists,0), V.AllowTransactions,  "&_
		'		 "V.EligibleForTDS, V.OUDefinitionID  FROM dbo.VwOrgGLHeads V INNER JOIN "&_
		'		 "dbo.GetGLOpen1('"&sFinMonYear&"', '"&sCurrDate&"', '0', '"&sOrgID&"') F ON V.AccountHead = F.AccHD "&_
		'		 "WHERE V.SubLedger = '0' and V.OUDefinitionID = '"&sOrgID&"' and AccountHead <> '"&iBkAccHead&"' ORDER BY 1 "

		sQuery = "Select V.AccountDescription,isNull(dbo.GetGLOpen('"&sFinMonYear&"','"&sCurrDate&"','0','"&sOrgID&"',V.AccountHead),'0.00') Balance, "&_
				 "V.AccountHead,V.AccountHeadCode,isnull(V.CostCenterExists,'0'), "&_
				 "isnull(V.AnalyticalHeadExists,'0'),V.AllowTransactions,V.EligibleForTds "&_
				 "From VwOrgGLHeads V Where V.OUDefinitionID = '"&sOrgID&"' And V.SubLedger = '0' and AccountHead <> '"&iBkAccHead&"' Order BY 1 "
		

	End IF
else
	IF CStr(iBkAccHead) = "" Then
		'oDataList.sSQL = "select AccountDescription,AccountHead,AccountHeadCode,isnull(CostCenterExists,'0'),"&_
		'				 "isnull(AnalyticalHeadExists,'0'),AllowTransactions,EligibleForTds from VwOrgGLHeads where SubLedger=0 and OUDefinitionID='"&sOrgId&"' "&_
		'				 " and AccountDescription Like '"&sFilter&"' order by AccountDescription "

		'sQuery = "SELECT V.AccountDescription, isNull(F.CloseAmtWithIndi,0.00),V.AccountHead, V.AccountHeadCode, "&_
		'		 "isNull(V.CostCenterExists,0),isNull(V.AnalyticalHeadExists,0), V.AllowTransactions,  "&_
		'		 "V.EligibleForTDS, V.OUDefinitionID  FROM dbo.VwOrgGLHeads V INNER JOIN "&_
		'		 "dbo.GetGLOpen1('"&sFinMonYear&"', '10/03/2006', '0', '"&sOrgID&"') F ON V.AccountHead = F.AccHD "&_
		'		 "WHERE V.SubLedger = '0' and V.OUDefinitionID = '"&sOrgID&"'   "&_
		'		 "and AccountDescription Like '"&sFilter&"' Order BY 1 "


		sQuery = "Select V.AccountDescription,isNull(dbo.GetGLOpen('"&sFinMonYear&"','"&sCurrDate&"','0','"&sOrgID&"',V.AccountHead),'0.00') Balance, "&_
				 "V.AccountHead,V.AccountHeadCode,isnull(V.CostCenterExists,'0'), "&_
				 "isnull(V.AnalyticalHeadExists,'0'),V.AllowTransactions,V.EligibleForTds "&_
				 "From VwOrgGLHeads V Where V.OUDefinitionID = '"&sOrgID&"' And V.SubLedger = '0'  " 
				 '"And V.AccountDescription Like '"&sFilter&"' Order BY 1 "
		if Trim(sSearchBy) = "IN" then
			sFilter =  sFilter & "%"
			sQuery = sQuery & " And V.AccountDescription like " & Pack(sFilter) & " "
		elseif Trim(sSearchBy) = "IA" then	
			 sFilter = "%" & sFilter
			sQuery = sQuery & " And V.AccountDescription like " & Pack(sFilter) & " "
		end if
		sQuery = sQuery & " Order BY 1 "


	Else
		'oDataList.sSQL = "select AccountDescription,AccountHead,AccountHeadCode,isnull(CostCenterExists,'0'),"&_
		'				 "isnull(AnalyticalHeadExists,'0'),AllowTransactions,EligibleForTds from VwOrgGLHeads where SubLedger=0 and OUDefinitionID='"&sOrgId&"' and AccountHead <> '"&iBkAccHead&"' "&_
		'				 " and AccountDescription Like '"&sFilter&"' order by AccountDescription "

		'sQuery = "SELECT V.AccountDescription, isNull(F.CloseAmtWithIndi,0.00),V.AccountHead, V.AccountHeadCode, "&_
		'		 "isNull(V.CostCenterExists,0),isNull(V.AnalyticalHeadExists,0), V.AllowTransactions,  "&_
		'		 "V.EligibleForTDS, V.OUDefinitionID  FROM dbo.VwOrgGLHeads V INNER JOIN "&_
		'		 "dbo.GetGLOpen1('"&sFinMonYear&"', '10/03/2006', '0', '"&sOrgID&"') F ON V.AccountHead = F.AccHD "&_
		'		 "WHERE V.SubLedger = '0' and V.OUDefinitionID = '"&sOrgID&"' and AccountHead <> '"&iBkAccHead&"'  "&_
		'		 "and AccountDescription Like '"&sFilter&"' Order BY 1 "

		sQuery = "Select V.AccountDescription,isNull(dbo.GetGLOpen('"&sFinMonYear&"','"&sCurrDate&"','0','"&sOrgID&"',V.AccountHead),'0.00') Balance, "&_
				 "V.AccountHead,V.AccountHeadCode,isnull(V.CostCenterExists,'0'), "&_
				 "isnull(V.AnalyticalHeadExists,'0'),V.AllowTransactions,V.EligibleForTds "&_
				 "From VwOrgGLHeads V Where V.OUDefinitionID = '"&sOrgID&"' And V.SubLedger = '0' and  "&_
				 "AccountHead <> '"&iBkAccHead&"' "
				  ''''''' and AccountDescription Like '"&sFilter&"' Order BY 1 "
		if Trim(sSearchBy) = "IN" then
			sFilter =  sFilter & "%"
			sQuery = sQuery & " And V.AccountDescription like " & Pack(sFilter) & " "
		elseif Trim(sSearchBy) = "IA" then				
			sFilter = "%" & sFilter 
			sQuery = sQuery & " And V.AccountDescription like " & Pack(sFilter) & " "
		end if
		sQuery = sQuery & " Order BY 1 "


	End IF
end if

'Response.Write sQuery
oDatalist.sSql = sQuery
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>
</BODY>
</HTML>