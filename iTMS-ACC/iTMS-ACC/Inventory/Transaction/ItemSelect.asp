<%
	'Program Name				:	ITEMSELECT.asp
	'Module Name				:	To List all Item details
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	APRIL 05,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	REQFORQUOTELIST.ASP
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
<!-- Conversion note: NewclsItemDataListMultiple.asp was not available; using clsItemDataListMultiple.asp for testing. -->
<!--#INCLUDE FILE="../../include/clsItemDataListMultiple.asp"-->
<!--#include file="../../include/populate.asp"-->
<HTML><HEAD><TITLE>Item</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable
dim sIType,sOrgID,sFilter,sSearchBy,sSelectMode,sFlag
dim iStock

'Response.Write "<p><p>Request.QueryString = " & Request.QueryString

sOrgID = trim(Request.QueryString("orgID"))
sIType = trim(Request.QueryString("sIType"))
sFilter = trim(Request.QueryString("Query"))
sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
sFlag = trim(Request.QueryString("Flag"))
iStock =  trim(Request.QueryString("Stock"))
'Response.Write "iStock="&iStock
if sIType = "" then
	sIType = "STO"
else
	sIType = right(sIType,3)
end if

if trim(sSelectMode) = "" then sSelectMode = "R"

'Response.Write "<p>sSelectMode = " & sSelectMode
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


oDataList.AddEnableField sFlag
oDataList.AddOptDispField sSelectMode
oDataList.AddReturnedField "0" 'Companyitemcode
oDataList.AddReturnedField "3"	'ItemCode
oDataList.AddReturnedField "4" ' class code
oDataList.AddReturnedField "2"  'group name
oDataList.AddReturnedField "1" ' item name
oDataList.AddReturnedField "5" ' stores uom
oDataList.AddReturnedField "6" ' Decimal
oDataList.AddReturnedField "7" ' Receipt Numbering

oDataList.SearchForDesc "Search For"
oDataList.AddSearchField "Item Code","IC"
oDataList.AddSearchField "Item Name < Starts with >","IN"
oDataList.AddSearchField "Item Name < Anywhere >","IA"
oDataList.AddSearchField "Drawing Number","DN"
oDataList.AddSearchField "Classification","CL"

'oDataList.SearchForDesc "Item Description"

'response.write sSearchBy & "<BR>"

' Declaration of variables
Dim dcrs,stypID,stypName

iSAApplicationPop = Session("iApplication")
iSAProcessPop = Session("iProcess")
iSAActivityPop = Session("iActivity")
iEmpNoPopulate = Session("employeenumber")

'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	if iSAApplicationPop <> "" then
	.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE WHERE ITEMTYPEID IN (SELECT DISTINCT ITEMTYPEID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY ITEMTYPENO"
	else
	.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE ORDER BY ITEMTYPENO"
	end if
	.ActiveConnection = con
	.Open
end with

set dcrs.ActiveConnection = nothing
set stypID = dcrs(0)
set stypName = dcrs(1)

If not dcrs.EOF then
	Do While Not dcrs.EOF
		oDataList.AddSearchField1 trim(stypName),trim(stypID)
	dcrs.MoveNext
	Loop
End If
dcrs.Close

'if ucase(trim(sIType)) = "STO" then

if 1 = 1 then
	if sFilter = "" then
		oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ltrim(rtrim(ITEMDESCRIPTION)) + ' -- ' + ltrim(rtrim(ISNULL(DRAWINGNUMBER,'N/A')))  + ' -- ' + ltrim(rtrim(ISNULL(CATALOGUENO,'N/A'))) ,GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM,DECIMALALLOWED,RECEIPTNUMBERING FROM VWITEM WHERE ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
	elseif sFilter <> "" then
		if Trim(sSearchBy) = "IC" then
			sFilter = "%" & sFilter & "%"
			oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ltrim(rtrim(ITEMDESCRIPTION)) + ' -- ' + ltrim(rtrim(ISNULL(DRAWINGNUMBER,'N/A')))  + ' -- ' + ltrim(rtrim(ISNULL(CATALOGUENO,'N/A'))),GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM,DECIMALALLOWED,RECEIPTNUMBERING FROM VWITEM WHERE COMPANYITEMCODE like " & Pack(trim(sFilter)) & " AND ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
		elseif Trim(sSearchBy) = "IN" then
			sFilter =  sFilter & "%"
			oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ltrim(rtrim(ITEMDESCRIPTION)) + ' -- ' + ltrim(rtrim(ISNULL(DRAWINGNUMBER,'N/A')))  + ' -- ' + ltrim(rtrim(ISNULL(CATALOGUENO,'N/A'))),GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM,DECIMALALLOWED,RECEIPTNUMBERING FROM VWITEM WHERE ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
		elseif Trim(sSearchBy) = "IA" then
			sFilter = "%" & sFilter & "%"
			oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ltrim(rtrim(ITEMDESCRIPTION)) + ' -- ' + ltrim(rtrim(ISNULL(DRAWINGNUMBER,'N/A')))  + ' -- ' + ltrim(rtrim(ISNULL(CATALOGUENO,'N/A'))),GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM,DECIMALALLOWED,RECEIPTNUMBERING FROM VWITEM WHERE ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
		elseif Trim(sSearchBy) = "DN" then
			sFilter = "%"&sFilter&"%"
			oDataList.sSQL = "SELECT DISTINCT IM.COMPANYITEMCODE,ltrim(rtrim(IM.ITEMDESCRIPTION)) + ' -- ' + ltrim(rtrim(ISNULL(IM.DRAWINGNUMBER,'N/A')))  + ' -- ' + ltrim(rtrim(ISNULL(IM.CATALOGUENO,'N/A'))),IM.ITEMCODE FROM INV_M_ITEMMASTER IM,INV_M_CLASSIFICATION IC,INV_M_ITEMGROUP IG WHERE IM.DRAWINGNUMBER LIKE " & Pack(sFilter)& " AND IC.ITEMTYPEID = " & Pack(sIType) & " AND IC.GROUPCODE = IG.CLASSIFICATIONCODE AND IG.LEAFNODE = 1 AND IM.ITEMCODE = IG.ITEMCODE AND STR(IC.GROUPCODE)+STR(IM.ITEMCODE) NOT IN (SELECT STR(CLASSIFICATIONCODE)+STR(ITEMCODE) FROM INV_M_ITEMORGMASTER WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0) ORDER BY 2"
		elseif Trim(sSearchBy) = "CL" then
			sFilter = "%" & sFilter & "%"
			oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ltrim(rtrim(ITEMDESCRIPTION)) + ' -- ' + ltrim(rtrim(ISNULL(DRAWINGNUMBER,'N/A')))  + ' -- ' + ltrim(rtrim(ISNULL(CATALOGUENO,'N/A'))),GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM,DECIMALALLOWED,RECEIPTNUMBERING FROM VWITEM WHERE GROUPNAME like " & Pack(trim(sFilter)) & " AND ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
		end if
	end if

else

	if sFilter = "" then
		oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM FROM VWITEM WHERE ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
	elseif sFilter <> "" then
		if Trim(sSearchBy) = "IC" then
			sFilter = "%" & sFilter & "%"
			oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM FROM VWITEM WHERE COMPANYITEMCODE like " & Pack(trim(sFilter)) & " AND ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
		elseif Trim(sSearchBy) = "IN" then
			sFilter =  sFilter & "%"
			oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM FROM VWITEM WHERE ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
		elseif Trim(sSearchBy) = "IA" then
			sFilter = "%" & sFilter & "%"
			oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM FROM VWITEM WHERE ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
		elseif Trim(sSearchBy) = "DN" then
			sFilter = "%"&sFilter&"%"
			oDataList.sSQL = "SELECT DISTINCT IM.COMPANYITEMCODE,IM.ITEMDESCRIPTION,IM.ITEMCODE FROM INV_M_ITEMMASTER IM,INV_M_CLASSIFICATION IC,INV_M_ITEMGROUP IG WHERE IM.DRAWINGNUMBER LIKE " & Pack(sFilter)& " AND IC.ITEMTYPEID = " & Pack(sIType) & " AND IC.GROUPCODE = IG.CLASSIFICATIONCODE AND IG.LEAFNODE = 1 AND IM.ITEMCODE = IG.ITEMCODE AND STR(IC.GROUPCODE)+STR(IM.ITEMCODE) NOT IN (SELECT STR(CLASSIFICATIONCODE)+STR(ITEMCODE) FROM INV_M_ITEMMASTER WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0) ORDER BY 2"
		elseif Trim(sSearchBy) = "CL" then
			sFilter = "%" & sFilter & "%"
			oDataList.sSQL = "SELECT DISTINCT COMPANYITEMCODE,ITEMDESCRIPTION,GROUPNAME,ITEMCODE,CLASSIFICATIONCODE,STORESUOM FROM VWITEM WHERE GROUPNAME like " & Pack(trim(sFilter)) & " AND ITEMTYPEID = " & Pack(sIType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND PURCHASEELIGIBLE = 1 AND ITEMACTIVE = 'Y' AND ITEMONHOLD = 0 ORDER BY 2"
		end if
	end if 'if sFilter = "" then
end if 'if ucase(trim(sIType)) = "STO" then

'Response.Write " <p> oDataList.sSQL = " & oDataList.sSQL
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
sFilter = ""
%>
</BODY>
</HTML>

