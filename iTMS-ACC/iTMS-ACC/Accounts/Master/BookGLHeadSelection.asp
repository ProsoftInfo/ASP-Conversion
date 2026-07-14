<%Option Explicit%>
<!--#include virtual="/include/clsDatabase.asp"-->
<!--#include virtual="/include/clsGLDatalist.asp"-->
<!--#include virtual="/include/populate.asp"-->
<HTML><HEAD><TITLE>GL Head Selection</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sQuery
dim sIType,sOrgID,sFilter,sBookCode,iBookNo,iBkAccHead

sOrgID = trim(Request.QueryString("orgID"))
sBookCode = trim(Request.QueryString("BookCode"))
sFilter = trim(Request.QueryString("Query"))&"%"

'Response.Write sOrgID

Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist

oDataList.PageSize = 15
oDataList.PrimaryKey = "AccountHead"

'oDataList.AddDisplayField "Item Code"
oDataList.AddDisplayField "Account Head Description"

oDataList.AddReturnedField "1"
oDataList.AddReturnedField "0"
oDataList.AddReturnedField "2"
oDataList.AddReturnedField "0"
oDataList.AddReturnedField "0"
oDataList.AddReturnedField "0"

'Response.Write sFilter & sBookCode & sBookCode

if sFilter = "%" and CStr(sBookCode) = "04" then
	oDataList.sSQL = "select Distinct AccountDescription,AccountHead,AccountHeadCode from VwOrgGLHeads where AccountHead "&_
					 "not in (select BookAccountHead from Acc_R_ApplicableAccountHeads where BookAccountHead is not null and OUDefinitionID='"&sOrgId&"' and Useable <> '1' )"&_
					 " and SubLedger=0 and OUDefinitionID='"&sOrgId&"' order by AccountDescription"
elseif sFilter <> "%" and CStr(sBookCode) = "04" then
	oDataList.sSQL = "select Distinct AccountDescription,AccountHead,AccountHeadCode from VwOrgGLHeads where AccountHead "&_
					 "not in (select BookAccountHead from Acc_R_ApplicableAccountHeads where BookAccountHead is not null and OUDefinitionID='"&sOrgId&"' and Useable <> '1')"&_
					 " and SubLedger=0 and OUDefinitionID='"&sOrgId&"' and AccountDescription Like '"&sFilter&"' order by AccountDescription"
elseif sFilter = "%" and CStr(sBookCode) = "05" then
	oDataList.sSQL = "Select Distinct M.AccountDescription,M.AccountHead,M.AccountHead From  "&_
					 "Acc_M_GLAccountHead M,APP_R_OrgnTaxAccountHead R WHere "&_
					 "R.AccountHead = M.AccountHead and R.OUDefinitionID = '"&sOrgID&"' and R.TaxCode  "&_
					 "is NULL Order By M.AccountDescription "
elseif sFilter <> "%" and CStr(sBookCode) = "05" then
	oDataList.sSQL = "Select Distinct M.AccountDescription,M.AccountHead,M.AccountHead From  "&_
					 "Acc_M_GLAccountHead M,APP_R_OrgnTaxAccountHead R WHere "&_
					 "R.AccountHead = M.AccountHead and R.OUDefinitionID = '"&sOrgID&"' and R.TaxCode  "&_
					 "is NULL and M.AccountDescription Like '"&sFilter&"' Order By M.AccountDescription "
Elseif sFilter = "%" and CStr(sBookCode) <> "04" and CStr(sBookCode) <> "05" Then
	oDataList.sSQL = "select Distinct AccountDescription,AccountHead,AccountHeadCode from VwOrgGLHeads where AccountHead "&_
					 "not in (select BookAccountHead from Acc_R_ApplicableAccountHeads where BookAccountHead is not null and OUDefinitionID='"&sOrgId&"' and Useable='0')"&_
					 " and SubLedger=0 and OUDefinitionID='"&sOrgId&"' order by AccountDescription"
Elseif sFilter <> "%" and CStr(sBookCode) <> "04" and CStr(sBookCode) <> "05" Then
	oDataList.sSQL = "select Distinct AccountDescription,AccountHead,AccountHeadCode from VwOrgGLHeads where AccountHead "&_
					 "not in (select BookAccountHead from Acc_R_ApplicableAccountHeads where BookAccountHead is not null and OUDefinitionID='"&sOrgId&"' and Useable='0')"&_
					 " and SubLedger=0 and OUDefinitionID='"&sOrgId&"' and AccountDescription Like '"&sFilter&"' order by AccountDescription"
end if
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>
</BODY>
</HTML>