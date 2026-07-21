<%
	'Program Name				:	ItemSelectForReport.asp
	'Module Name				:
	'Author Name				:	R. Ragavendran
	'Created On					:	August 17,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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


<%Option Explicit%>
<!--#INCLUDE FILE="../../include/clsDatabase.asp"-->
<!--#INCLUDE FILE="../../include/clsItemDataListMultiple.asp"-->
<!--#include file="../../include/populate.asp"-->
<HTML><HEAD><TITLE>Item</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable,sTempSer
dim sIType,sOrgID,sFilter,sSearchBy,sSelectMode,sFinPeriod,sFinYearFrom,sFinYearTo,sQuery


'Response.Write "<p><p>Request.QueryString = " & Request.QueryString
sSearchBy = ""

Response.Write sSearchBy &"<br>"
'sOrgID = trim(Request.QueryString("orgID"))
sOrgID = Session("organizationcode")
sIType = trim(Request.QueryString("iType"))
sFilter = trim(Request.QueryString("Query"))
'MsgBox sFilter
sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
sTempSer = Split(trim(Request.QueryString("SearchBy")),",")

'Needs to add Financial year filter to avoid duplicate entry
sFinPeriod = split(trim(Session("FinPeriod")),":")
sFinYearFrom = "01/04/"&sFinPeriod(0)
sFinYearTo = "31/03/"&sFinPeriod(1)
Response.Write "<font color=red>"

IF UBound(sTempSer) > 0 Then
	sSearchBy = Trim(sTempSer(0))
Else
	sSearchBy = trim(Request.QueryString("SearchBy"))
End IF
if trim(sIType)="" or IsNull(sIType) then sIType = ""

'Response.Write "<p>Search By Val " &  sSearchBy
'Response.Write UBound(sTempSer)
'Response.Write "<p>sOrgID " & Request.QueryString("orgID")
'Response.Write "<p>sIType " & Request.QueryString("iType")
'Response.Write "<p>sFilter " & Request.QueryString("Query")
'Response.Write "<p>sSearchBy " & Request.QueryString("SearchBy")
'Response.Write "<p>sSelectMode " & Request.QueryString("hSelectMode")

if trim(sSelectMode) = "" then sSelectMode = "R"

'Response.Write "<p>sSearchBy = " & sSearchBy
Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist

oDataList.PageSize = 15

oDataList.PrimaryKey = "ItemCode"

oDataList.AddDisplayField "Item Code"
if ucase(trim(sIType)) = "STO" then
	oDataList.AddDisplayField "Item Desc - Dr No - Ca No"
else
	oDataList.AddDisplayField "Item Description"
end if
oDataList.AddDisplayField "Classification"
oDataList.AddDisplayField "Stock"
oDataList.AddDisplayField "UOM"
'VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFI
oDataList.AddOptDispField sSelectMode


oDataList.AddReturnedField "0" 'Companyitemcode
oDataList.AddReturnedField "5"	'ItemCode
oDataList.AddReturnedField "6" ' class code
oDataList.AddReturnedField "2"  'group name
oDataList.AddReturnedField "1" ' item name
oDataList.AddReturnedField "4"	'Stores UOM


oDataList.SearchForDesc "Search For"
oDataList.AddSearchField "Item Code","IC"
oDataList.AddSearchField "Item Name < Starts with >","IN"
oDataList.AddSearchField "Item Name < Anywhere >","IA"
oDataList.AddSearchField "Drawing Number","DN"
oDataList.AddSearchField "Classification","CL"
oDataList.AddSearchField "Catalogue No.","CN"
oDataList.AddSearchField "MGR No.","MN"
oDataList.AddSearchField "Page No.","PN"
oDataList.AddSearchField "Position No.","PO"
'oDataList.AddSearchField "Work Center Name","WC"
'oDataList.AddSearchField "Machine Center Name","MC"

if sFilter = "" then
	sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE FROM VWITEM VI,VwYearlyStock VS "
	sQuery = sQuery & " Where VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.PURCHASEELIGIBLE = 1 AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
	sQuery = sQuery & " ORDER BY 2"
elseif sFilter <> "" then
	' Item Codewise
	if Trim(sSearchBy) = "IC" then
		sFilter = "%" & sFilter & "%"
		sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DRAWINGNUMBER,VI.CATALOGUENO FROM VWITEM VI,VwYearlyStock VS WHERE VI.COMPANYITEMCODE like " & Pack(trim(sFilter)) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
		sQuery = sQuery & " ORDER BY 2"
	' Item Namewise Starts with
	elseif Trim(sSearchBy) = "IN" then
		sFilter =  sFilter & "%"
		sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DRAWINGNUMBER,VI.CATALOGUENO FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND  CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
		 sQuery = sQuery & " ORDER BY 2"
	' Item Namewise Anywhere
	elseif Trim(sSearchBy) = "IA" then
		sFilter = "%" & sFilter & "%"
		sQuery= "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DRAWINGNUMBER,VI.CATALOGUENO FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103) "
		 sQuery = sQuery &  " ORDER BY 2"
	' Item Drawing Numberwise
	elseif Trim(sSearchBy) = "DN" then
		sFilter = "%"&sFilter&"%"
		sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DRAWINGNUMBER,VI.CATALOGUENO FROM VWITEM VI,VwYearlyStock VS WHERE VI.DRAWINGNUMBER LIKE " & Pack(sFilter)& "  AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103) "
		 sQuery = sQuery & " ORDER BY 2"
	' Classification Namewise
	elseif Trim(sSearchBy) = "CL" then
		sFilter = "%" & sFilter & "%"
		sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DRAWINGNUMBER,VI.CATALOGUENO FROM VWITEM VI,VwYearlyStock VS WHERE VI.GROUPNAME like " & Pack(trim(sFilter)) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103) "
		 sQuery = sQuery & " ORDER BY 2"
	' Item Catalogue Numberwise
	elseif Trim(sSearchBy) = "CN" then
		sFilter =  sFilter & "%"
		sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DRAWINGNUMBER,VI.CATALOGUENO FROM VWITEM VI,VwYearlyStock VS WHERE VI.CATALOGUENO LIKE " & Pack(sFilter) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
		 sQuery = sQuery & " ORDER BY 2"
	' MGR Numberwise
	elseif Trim(sSearchBy) = "MN" then
		sFilter =  sFilter & "%"
		sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DRAWINGNUMBER,VI.CATALOGUENO FROM VWITEM VI,VwYearlyStock VS WHERE VI.MGRNO LIKE " & Pack(sFilter) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
		 sQuery = sQuery & " ORDER BY 2"
	' Page Numberwise
	elseif Trim(sSearchBy) = "PN" then
		sFilter =  sFilter & "%"
		sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DRAWINGNUMBER,VI.CATALOGUENO FROM VWITEM VI,VwYearlyStock VS WHERE VI.PAGENO LIKE " & Pack(sFilter) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
		 sQuery = sQuery & " ORDER BY 2"
	' Position Numberwise
	elseif Trim(sSearchBy) = "PO" then
		sFilter =  sFilter & "%"
		sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,(VI.ITEMDESCRIPTION+'-'+isNull(VI.DRAWINGNUMBER,'')+'-'+isnull(VI.CATALOGUENO,'')),VI.GROUPNAME,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DRAWINGNUMBER,VI.CATALOGUENO FROM VWITEM VI,VwYearlyStock VS WHERE VI.POSITIONNO LIKE " & Pack(sFilter) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND  CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103) "
		 sQuery = sQuery & " ORDER BY 2"
	end if
end if

oDataList.sSQL = sQuery

Response.Write " <p> oDataList.sSQL = " & oDataList.sSQL
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
sFilter = ""
%>
</BODY>
</HTML>

