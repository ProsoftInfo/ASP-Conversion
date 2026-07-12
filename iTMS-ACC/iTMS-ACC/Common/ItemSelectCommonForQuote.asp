 <%
	'Program Name				:	ItemSelectCommonForQuote.asp
	'Module Name				:	To List all Item details in Valid Quote
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 04,2012
	'Modified By				:
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
<!--#INCLUDE FILE="../include/clsDatabase.asp"-->
<!--#INCLUDE FILE="../include/clsItemDataListMultiple.asp"-->
<!--#include file="../include/populate.asp"-->
<HTML><HEAD><TITLE>Item</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable
Dim sIType,sOrgID,sFilter,sSearchBy,sSelectMode,sFlag,sQuery,sFlagItemStock
Dim sFinPeriod,sFinYearFrom,sFinYearTo,sTempMonYr,sMonYr,sButtDispMode,sPartyCode
Dim iStock,iClassCodes
Response.Write "<font color=red>"

'Response.Write "<p><p>Request.QueryString = " & Request.QueryString

sOrgID = trim(Request.QueryString("orgID"))
if Trim(sOrgID)="" then sOrgID = Session("organizationcode")
sIType = trim(Request.QueryString("sIType"))
sFilter = trim(Request.QueryString("Query"))
iStock =  trim(Request.QueryString("Stock")) 'Newly Added
sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
sFlag = trim(Request.QueryString("Flag"))
iClassCodes = Request.QueryString("hClassCodes")
sButtDispMode = UCase(trim(Request.QueryString("hDispButt")))
sFlagItemStock = UCase(trim(Request.QueryString("hDispItem")))
sPartyCode = Trim(Request.QueryString("hPartyCode"))

'Response.Write "sPartyCode = "& sPartyCode
'Response.Write sSelectMode
if len(Month(date())) = 1 then
	sTempMonYr = "0"&Month(date())
else
	sTempMonYr = Month(date())
end if
sMonYr = sTempMonYr&Year(date())

'sFinPeriod = split(GetFinancialYear(sMonYr),":")
'sFinYearFrom =  sFinPeriod(0)
'sFinYearTo = sFinPeriod(1)

''''IF Login Financial Year has to be used '''''''
'                                                '
	sFinPeriod = split(Session("FinPeriod"),":") '
	sFinYearFrom =  "01/04/"&sFinPeriod(0)       '
	sFinYearTo = "31/03/"&sFinPeriod(1)          '
'			                      				 '
''''''''''''''''''''''''''''''''''''''''''''''''''

'Response.Write "iStock="&iStock
if sIType = "" then
	sIType = "STO"
else
	sIType = right(sIType,3)
end if
'sIType = "GAR"
if trim(sSelectMode) = "" then sSelectMode = "R"
if trim(sButtDispMode)="" then sButtDispMode = "N"
if trim(sFlagItemStock)="" then sFlagItemStock = 0
'Response.Write "sIType="&sIType
'Response.Write "<p>sSelectMode = " & sSelectMode
Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist

oDataList.PageSize = 15
oDataList.PartyCode = sPartyCode
oDataList.PrimaryKey = "ItemCode"

oDataList.AddDisplayField "Item Code"
if ucase(trim(sIType)) = "STO" then
	oDataList.AddDisplayField "Item Description"
else
	oDataList.AddDisplayField "Item Description"
end if
oDataList.AddDisplayField "Classification"
IF trim(istock) = "Y" then
	oDataList.AddDisplayField "Stock"
	oDataList.AddDisplayField "UOM"
End IF
'if sIType = "GAR" or sIType = "FAB" or sIType = "FIB"  then
	oDataList.AddDisplayField "AttributeList"
'end if
'oDataList.AddEnableField sFlag
oDataList.AddOptDispField sSelectMode
oDatalist.AddDispAddButt sButtDispMode
oDatalist.AddDispItem sFlagItemStock
oDataList.AddReturnedField "0" 'Companyitemcode
oDataList.AddReturnedField "5" ' "3"	'ItemCode
oDataList.AddReturnedField "6" '"4" ' class code
oDataList.AddReturnedField "2"  'group name
oDataList.AddReturnedField "1" ' item name
oDataList.AddReturnedField "4" '"5" ' stores uom
oDataList.AddReturnedField "7" '"6" ' Decimal
oDataList.AddReturnedField "8" '"7" ' Receipt Numbering
oDataList.AddReturnedField "9" '"8" ' AttributeList

oDataList.SearchForDesc "Search For"
oDataList.AddSearchField "Item Code","IC"
oDataList.AddSearchField "Item Name < Starts with >","IN"
oDataList.AddSearchField "Item Name < Anywhere >","IA"
oDataList.AddSearchField "Drawing Number","DN"

oDataList.AddSearchTypeField "Company","CO"
oDataList.AddSearchTypeField "Party","PA"

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
'with dcrs
'	.CursorLocation = 3
'	.CursorType = 3
'	.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE ORDER BY ITEMTYPENO"
'	.ActiveConnection = con
'	.Open
'end with
'Response.Write dcrs.Source
'set dcrs.ActiveConnection = nothing
'If not dcrs.EOF then
'	Do While Not dcrs.EOF
'	    stypID = dcrs(0)
'        stypName = dcrs(1)
'		oDataList.AddSearchField1 trim(stypName),trim(stypID)
'	dcrs.MoveNext
'	Loop
'End If
'dcrs.Close

'if ucase(trim(sIType)) = "STO" then
	if sFilter = "" then
	'	sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
     '      "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.PURCHASEELIGIBLE = 1 AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
           sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
           "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.PURCHASEELIGIBLE = 1 AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
	elseif sFilter <> "" then
		if Trim(sSearchBy) = "IC" then
			sFilter = "%" & sFilter & "%"
'			sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
 '               "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.COMPANYITEMCODE like " & Pack(trim(sFilter)) & " AND VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
            sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
                "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.COMPANYITEMCODE like " & Pack(trim(sFilter)) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
		elseif Trim(sSearchBy) = "IN" then
			sFilter =  sFilter & "%"
'			sQuery= "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
 '               "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
 			sQuery= "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
                "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND  VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"

		elseif Trim(sSearchBy) = "IA" then
			sFilter = "%" & sFilter & "%"
			'sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
            '    "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
            sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
                "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMDESCRIPTION LIKE " & Pack(sFilter) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
		elseif Trim(sSearchBy) = "DN" then
			sFilter = "%"&sFilter&"%"
		'	sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION ,(Select GroupName from INV_M_Classification "&_
         '                  "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.DRAWINGNUMBER LIKE " & Pack(sFilter)& " AND VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103) "
            sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION ,(Select GroupName from INV_M_Classification "&_
                           "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.DRAWINGNUMBER LIKE " & Pack(sFilter)& " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103) "
		end if
	end if

'else
'
'	if sFilter = "" then
'		sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
 '                  "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
'	elseif sFilter <> "" then
'		if Trim(sSearchBy) = "IC" then
'			sFilter = "%" & sFilter & "%"
'			sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
 '                   "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.COMPANYITEMCODE like " & Pack(trim(sFilter)) & " VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
'		elseif Trim(sSearchBy) = "IN" then
'			sFilter =  sFilter & "%"
'			sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
 '               "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMDESCRIPTION LIKE " & Pack(trim(sFilter)) & " VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
'		elseif Trim(sSearchBy) = "IA" then
'			sFilter = "%" & sFilter & "%"
'			sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
 '           "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ITEMDESCRIPTION LIKE " & Pack(trim(sFilter)) & " VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
'		elseif Trim(sSearchBy) = "DN" then
'			sFilter = "%"&sFilter&"%"
'			sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,VI.ITEMCODE,(Select GroupName from INV_M_Classification "&_
 '               "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.DRAWINGNUMBER LIKE " & Pack(sFilter)& " AND VI.ITEMTYPEID = " & Pack(sIType) & " AND VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND VI.ITEMTYPEID = VS.ITEMTYPEID AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
'		end if
'	end if 'if sFilter = "" then
'end if 'if ucase(trim(sIType)) = "STO" then

if trim(iClassCodes)<>"" then
    sQuery = sQuery & " and VI.ClassificationCode in ("& iClassCodes &")"
end if

sQuery = sQuery & " and VI.ItemCode in (Select ItemCode from PUR_T_QuoteDetails where ItemCode not in("&_
                   " Select ItemCode from PUR_T_QuoteComparisonDetails where ItemStatus not in ('C','R') Group By ItemCode) Group By ItemCode) "
sQuery = sQuery & " Order By 2"
oDataList.sSQL = sQuery

'Response.Write " <p> sFilter = "& sFilter

'Response.Write " <p> oDataList.sSQL = <textarea>" & oDataList.sSQL&"</textarea>"
sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
sFilter = ""
%>
</BODY>
</HTML>

