<%@language="VBScript"%>
<%Option Explicit%>
<%
	'Program Name				:	XMLGetItemSelectRel.asp
	'Module Name				:	To Diplay The Item Details based on Search Condition
	'Author Name				:	Ragavendran R
	'Created On					:	Dec 19,2011
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
<!--#include file="../include/DatabaseConnection.asp"-->
<!--#include file="../include/populate.asp"-->
<!-- #include File="../include/CommonFunctions.asp" -->
<%

Dim sTable,oDOM
Dim sIType,sOrgID,sFilter,sSearchBy,sSelectMode,sFlag,sQuery,sFlagItemStock
Dim sFinPeriod,sFinYearFrom,sFinYearTo,sTempMonYr,sMonYr,sButtDispMode,sPartyCode
Dim sSuppItemCode,sSuppItemDesc
Dim iStock,iClassCodes,iCounter,iPageSize
Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,iSNo
Dim dcrs,rsTemp,rsTemp1
Dim DisableBut,sCheckNoOfBinAndLoc,sFinYrFrom,sFinYrTo,sLocNo,sBinNo
Dim nGetItemRate,nGetMarketPrice,sRequest,sSearchType
Dim ndRoot,ndItem
Dim sSearchItemCode,sSearchItemName,sSearchPartyItemCode,sSearchPartyItemName,sCheckStart,sPartyType,sSearchStock
Dim sFSNFlag,sRelatedTo,sCallFrom,nNoOfSupplier,sRelPartyType,sOrgnPartyCode,sPartyName
Dim sSuppParCode,sSuppParType,sSuppParSubType,sArrEligible,sPurEli,sSalEli,sInvEli,sManEli
Dim sUserAccessMode,sUserCode,sItemType,sCap,sDisplay

Set dcrs = Server.CreateObject("ADODB.Recordset")
Set rsTemp = Server.CreateObject("ADODB.Recordset")
Set rsTemp1 = Server.CreateObject("ADODB.Recordset")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

set ndRoot = oDOM.createElement("Root")
oDOM.appendChild ndRoot

'Response.Write "<font color=red>"

'sOrgID = trim(Request.QueryString("orgID"))
sOrgID = Session("organizationcode")
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
sSearchType = Request.QueryString("SearchType")
sSearchItemCode = Request.QueryString("SICode")
sSearchItemName = Request.QueryString("SIName")
sSearchPartyItemCode = Request.QueryString("SPCode")
sSearchPartyItemName = Request.QueryString("SPName")
sCheckStart = Request.QueryString("CheckStart")
sPartyType = Request.QueryString("PartyType")
sSearchStock = Request.QueryString("hStock")
sRelatedTo = Request.QueryString("RelatedTo")
sCallFrom = Request.QueryString("CallFrom")
sArrEligible = Split(Request.QueryString("Eligible"),":")
sItemType = Request.QueryString("IType")
sCap = Request.QueryString("Cap")
sDisplay =  Request.QueryString("Disp")
if trim(Request.QueryString("Eligible"))<>"" then
    sPurEli = sArrEligible(0)
    sSalEli = sArrEligible(1)
    sInvEli = sArrEligible(2)
    sManEli = sArrEligible(3)
else
    sPurEli = "N"
    sSalEli = "N"
    sInvEli = "N"
    sManEli = "N"
end if
sUserAccessMode = Request.QueryString("UAM")
sUserCode = Request.QueryString("UCODE")


if trim(sUserAccessMode)="" or IsNull(sUserAccessMode) then  sUserAccessMode="I"
if trim(sUserCode)="" or IsNull(sUserCode) then sUserCode=""
if trim(sCap)="" or IsNull(sCap) then sCap="N"

iPageSize = Request.QueryString("PageSize")
if Trim(iPageSize)="" or IsNull(iPageSize) then iPageSize = 10

if Trim(sSearchType)="" or IsNull(sSearchType) then sSearchType = "C"
if Trim(sSearchBy)="" or IsNull(sSearchBy) then sSearchBy = "IC"

iCurrentPage=Request("Page")
if Trim(iCurrentPage)="" or IsNull(iCurrentPage) then iCurrentPage = 1
iCurrentPage = CInt(iCurrentPage)

DisableBut = "N"

'Response.Write "sPartyCode = "& sPartyCode
'Response.Write sSelectMode
if len(Month(date())) = 1 then
	sTempMonYr = "0"&Month(date())
else
	sTempMonYr = Month(date())
end if
sMonYr = sTempMonYr&Year(date())

sFinPeriod = split(Session("FinPeriod"),":") '
sFinYearFrom =  "01/04/"&sFinPeriod(0)       '
sFinYearTo = "31/03/"&sFinPeriod(1)          '

if trim(sSelectMode) = "" then sSelectMode = "R"
if trim(sButtDispMode)="" then sButtDispMode = "N"
if trim(sFlagItemStock)="" then sFlagItemStock = 0

iSAApplicationPop = Session("iApplication")
iSAProcessPop = Session("iProcess")
iSAActivityPop = Session("iActivity")
iEmpNoPopulate = Session("employeenumber")

'Response.Write "sButtDispMode  ="& sButtDispMode

'Declaration of Objects

'if Trim(sSearchItemCode)<>"" or Trim(sSearchItemName)<>"" or trim(sSearchPartyItemCode)<>"" or Trim(sSearchPartyItemName)<>"" or Trim(sSearchStock)<>"" then
'    iCurrentPage = 1
'end if

if UCase(Trim(sCallFrom))="SUP" or UCase(Trim(sCallFrom))="SUPLIST"  then
    sRelPartyType="CR"
elseif UCase(Trim(sCallFrom))="CUS" or UCase(Trim(sCallFrom))="CUSLIST" then
    sRelPartyType="DR"
end if

sSearchItemCode = Request.QueryString("SICode")
sSearchItemName = Request.QueryString("SIName")
sSearchPartyItemCode = Request.QueryString("SPCode")
sSearchPartyItemName = Request.QueryString("SPName")

'Response.Write  vbCrLf & "sSearchPartyItemCode = "& sSearchPartyItemCode  & vbCrLf 
'Response.Write  vbCrLf & "sSearchPartyItemName = "& sSearchPartyItemName  & vbCrLf 
sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
           "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0),isNull(VS.YEARCLOSINGVALUE,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
           
if Trim(sDisplay)<>"ALL" then
        sQuery  =sQuery &" AND VI.ITEMACTIVE = 'Y'"
end if

if Trim(sSearchPartyItemCode)<>"" or trim(sSearchPartyItemName)<>"" or ucase(Trim(sCallFrom))="SUPLIST" or ucase(Trim(sCallFrom))="CUSLIST" then 
  sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM," &_
			 "VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0),isNull(VS.YEARCLOSINGVALUE,0),VI.PartyCode,VI.PartyType,VI.PartySubType FROM VW_ITEM_PartyDetails VI,VwYearlyStock VS " &_
             " WHERE VI.ORGANISATIONCODE = " & Pack(sOrgID) & " AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
             
             if Trim(sDisplay)<>"ALL" then
                sQuery  =sQuery &" AND VI.ITEMACTIVE = 'Y'"
             end if
             
             if Trim(sPartyCode)<>"" then
                sQuery = sQuery & "  and VI.PartyCode ="& sPartyCode
             end if
             
             if Trim(sSearchPartyItemCode)<>"" then
                sQuery = sQuery & " and SuppItemCode like '%" & sSearchPartyItemCode & "%'"
             end if
             
             if Trim(sSearchPartyItemName)<>"" then
                sQuery = sQuery & " and SuppItemDescription like "&pack("%"& sSearchPartyItemName &"%")
             end if
             
             if Trim(sSearchItemCode)<>"" then
                sQuery = sQuery & " and VI.COMPANYITEMCODE like "&Pack("%"& sSearchItemCode &"%")
             end if
             
             if Trim(sSearchItemName)<>"" then
                if sCheckStart="Y" then
                    sQuery = sQuery & " and VI.ITEMDESCRIPTION like "& Pack(sSearchItemName&"%")
                else
                    sQuery = sQuery & " and VI.ITEMDESCRIPTION like "&Pack("%"& sSearchItemName &"%")
                end if 
             end if
             
             if Trim(sSearchStock)<>"" then
                if InStr(1,sSearchStock,">")=1 or InStr(1,sSearchStock,"<")=1 or InStr(1,sSearchStock,"=")=1 then
                    sQuery = sQuery & " and VS.YEARCLOSINGSTOCK "& sSearchStock 
                else
                    sQuery = sQuery & " and VS.YEARCLOSINGSTOCK  = "& sSearchStock 
                end if
             end if
             
elseif (Trim(sSearchItemCode)<>"" or Trim(sSearchItemName)<>"") and Trim(sSearchPartyItemCode)="" and trim(sSearchPartyItemName) ="" or Trim(sSearchStock)<>"" then
    sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification "&_
               "where GroupCode in (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE,VI.DECIMALALLOWED,VI.RECEIPTNUMBERING,isNull(VI.ATTRIBUTELIST,0),isNull(VS.YEARCLOSINGVALUE,0) FROM VWITEM VI,VwYearlyStock VS WHERE VI.ORGANISATIONCODE = " & Pack(sOrgID) & "  AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) = CONVERT(Datetime,'"&sFinYearFrom&"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"&sFinYearTo&"',103)"
               
               
             if Trim(sDisplay)<>"ALL" then
                sQuery  =sQuery &" AND VI.ITEMACTIVE = 'Y'"
             end if
               
             if Trim(sSearchItemCode)<>"" then
                sQuery = sQuery & " and VI.COMPANYITEMCODE like '%"& sSearchItemCode &"%'"
             end if
             
             if Trim(sSearchItemName)<>"" then
                if sCheckStart="Y" then
                    sQuery = sQuery & " and VI.ITEMDESCRIPTION like "& Pack(sSearchItemName &"%")
                else
                    sQuery = sQuery & " and VI.ITEMDESCRIPTION like "& Pack("%"& sSearchItemName &"%")
                end if 
             end if
             
             if Trim(sSearchStock)<>"" then
                if InStr(1,sSearchStock,">")=1 or InStr(1,sSearchStock,"<")=1 or InStr(1,sSearchStock,"=")=1 then
                    sQuery = sQuery & " and VS.YEARCLOSINGSTOCK "& sSearchStock 
                else
                    sQuery = sQuery & " and VS.YEARCLOSINGSTOCK  = "& sSearchStock 
                end if
             end if
             
end if 'if (Trim(sSearchItemCode)<>"" or Trim(sSearchItemName)<>"") and Trim(sSearchPartyItemCode)="" and trim(sSearchPartyItemName) ="" then

if trim(iClassCodes)<>"" then
    sQuery = sQuery & " and VI.ClassificationCode in ("& iClassCodes &")"
end if

if Trim(sRelatedTo)<>"" then
    if trim(sRelatedTo) = "Y" then														
	    sQuery = sQuery & " and ltrim(VI.ItemCode)+':'+ltrim(VI.ClassificationCode) in ( Select ltrim(ItemCode)+':'+ltrim(ClassificationCode) from Inv_R_ItemSupplier where PartyType = '"& sRelPartyType &"')"
    elseif trim(sRelatedTo) = "N" then														
		sQuery = sQuery & " and ltrim(VI.ItemCode)+':'+ltrim(VI.ClassificationCode) NOT in ( Select ltrim(ItemCode)+':'+ltrim(ClassificationCode) from Inv_R_ItemSupplier where PartyType = '"& sRelPartyType &"')"	
	end if 	
end if

if trim(sPurEli)<>"N" then
    sQuery = sQuery & " and VI.PurchaseEligible = 1"
end if

if trim(sSalEli)<>"N" then
    sQuery = sQuery & " and VI.SalesEligible = 1"
end if

if trim(sInvEli)<>"N" then
    sQuery = sQuery & " and VI.InventoryEligible = 1"
end if

if trim(sManEli)<>"N" then
    sQuery = sQuery & " and VI.ManufactureEligible = 1"
end if

if trim(sUserAccessMode)="E" then
    if trim(sUserCode)<>"" then
        sQuery =  sQuery &" and VI.ItemCode in (Select ItemCode From Inv_R_ItemSupplier Where PartyCode = "& sUserCode &")"
    end if
end if

if UCase(trim(sItemType))<>"SELECT" and trim(sItemType)<>"" then
    sQuery = sQuery &" and ItemType = "&  sItemType
end if

if UCase(trim(sCap))="N" then
    sQuery =  sQuery &" and StockNonStock ='S'"
elseif uCase(trim(sCap))="C" then
    sQuery =  sQuery &" and StockNonStock ='N'"
end if

sQuery = sQuery & " Order By 2"
'Response.Clear 
'Response.Write vbCrLf & sQuery & vbCrLf
'Response.end
'Response.Write  "<textarea>"& sQuery &"</textarea>"

'Response.Write "<font color=red>"
    with dcrs
        .CursorLocation = 3
        .CursorType = 3
        .Source = sQuery
        .ActiveConnection = con
        .Open
    end with
    if not dcrs.EOF then
        iSNo = 1
        dcrs.PageSize=iPageSize
        if iCurrentPage=0 then iCurrentPage=1
        dcrs.AbsolutePage=iCurrentPage
        iTotalPage=dcrs.PageCount
        if cdbl(iCurrentPage) > cdbl(iTotalPage) then
            dcrs.AbsolutePage=1
        end if
        
        ndRoot.setAttribute "CurrPage",dcrs.AbsolutePage 
        ndRoot.setAttribute "TotPage",iTotalPage
        
        do while not dcrs.EOF and iSNo <=dcrs.PageSize 
        sSuppItemCode =""
        sSuppItemDesc =""
            if Trim(sPartyCode)<>"" then
                sQuery = "Select SuppItemCode,SuppItemDescription,PartyCode,PartyType,PartySubType from INV_R_ItemSupplier where PartyCode ="& sPartyCode &" and ItemCode ="&dcrs(5)&" and ClassificationCode ="& dcrs(6) &" and OrganisationCode = "& sOrgID
                'Response.Write sQuery
                rsTemp.Open sQuery,con
                if not rsTemp.EOF then
                    sSuppItemCode = rsTemp(0)
                    sSuppItemDesc = rsTemp(1)
                    sSuppParCode = rsTemp(2)
                    sSuppParType = rsTemp(3)
                    sSuppParSubType = rsTemp(4)
                end if
                rsTemp.Close
            end if 'if Trim(sPartyCode)<>"" then

            sCheckNoOfBinAndLoc = "0"
            sQuery = "Select count(*),LocationNumber,isNull(BinNumber,0) From Inv_T_ItemLocationStock where ItemCode="& dcrs(5) &" and ClassificationCode="& dcrs(6)&" and convert(DateTime,FinancialYearFrom,103) >= Convert(DateTime,'"& sFinYrFrom &"',103) and convert(DateTime,FinancialYearTo,103) <= Convert(DateTime,'"& sFinYrTo&"',103) Group by LocationNumber,BinNumber"
            rsTemp.Open sQuery,con
            If Not rsTemp.EOF Then
                Do while Not rsTemp.EOF
                    sCheckNoOfBinAndLoc = sCheckNoOfBinAndLoc + 1
                    sLocNo = rsTemp(1)
                    sBinNo = rsTemp(2)
                rsTemp.MoveNext
                Loop
            End IF
            rsTemp.Close
            if Trim(sPartyType)="DR" then
                nGetItemRate = GetItemSalePrice(Session("organizationcode"),Date(),dcrs(6),dcrs(5),sPartyCode)
            else
                nGetItemRate = GetItemPurchasePrice(Session("organizationcode"),Date(),dcrs(6),dcrs(5),sPartyCode)
            end if
            nGetMarketPrice = GetMarketPrice(Session("organizationcode"),dcrs(6),dcrs(5))
            
            
        	sQuery = "Select FSNCategory = CASE FSNCategory When 'F' Then 'Fast Moving' When 'S' Then 'Slow Moving' When 'N' Then 'Non-moving' Else '-' End from Inv_M_ItemOrgInventory where OrganisationCode = '"& sOrgID &"' and ItemCode = " & trim(dcrs(5)) & " and ClassificationCode = " & trim(dcrs(6))
        	rsTemp.Open sQuery,con
			if not rsTemp.EOF then
				sFSNFlag = rsTemp(0)
			Else
				sFSNFlag = "-"
			end if
			rsTemp.Close 
			
			'Count of SupplierCode
			if ucase(Trim(sCallFrom))="SUP" or ucase(Trim(sCallFrom))="CUS" then
			    if ucase(Trim(sCallFrom))="SUP" then
			        sQuery = "Select Count(*) From Inv_R_ItemSupplier Where ItemCode="& dcrs(5)&" and ClassificationCode = "& dcrs(6) &" and PartyType='CR'"
			    elseif ucase(Trim(sCallFrom))="CUS" then
			        sQuery = "Select Count(*) From Inv_R_ItemSupplier Where ItemCode="& dcrs(5)&" and ClassificationCode = "& dcrs(6) &" and PartyType='DR'"
			    end if' if Trim(sCallFrom)="S" then
			    rsTemp1.Open sQuery,con
			    If Not rsTemp1.EOF Then
				    nNoOfSupplier = rsTemp1(0)
				else
				    nNoOfSupplier = "0"
			    End IF
			    rsTemp1.Close 
			else
			    nNoOfSupplier = "0"
			end if' if Trim(sCallFrom)="Sup" or Trim(sCallFrom)="Cus" then
			
			if Trim(sPartyCode)<>"" then
			    sQuery = "Select OrgnPartyCode,PartyName from APP_M_PartyMaster where PartyCode = "& sPartyCode 
			    rsTemp.Open sQuery,con
			    if not rsTemp.EOF then
			        sOrgnPartyCode = rsTemp(0)
			        sPartyName = rsTemp(1)
			    end if
			    rsTemp.Close 
			else
			    sOrgnPartyCode = ""
			    sPartyName = ""
			end if
            
            set ndItem = oDOM.createElement("Item")
                ndItem.setAttribute "SNO",iSNo
                ndItem.setAttribute "ComItemCode",dcrs(0)
                ndItem.setAttribute "ItemName",dcrs(1)
                ndItem.setAttribute "PartyItemCode",sSuppItemCode 
                ndItem.setAttribute "PartyItemDesc",sSuppItemDesc
                ndItem.setAttribute "ClassName",dcrs(2)
                ndItem.setAttribute "Stock",dcrs(3)
                ndItem.setAttribute "UOM",dcrs(4)
                ndItem.setAttribute "ItemCode",dcrs(5)
                ndItem.setAttribute "ClassCode",dcrs(6)
                ndItem.setAttribute "DecimalAllowed",dcrs(7)
                ndItem.setAttribute "ReceiptNum",dcrs(8)
                ndItem.setAttribute "AttributeList",dcrs(9)
                ndItem.setAttribute "PartyCode",sPartyCode
                ndItem.setAttribute "ItemRate",nGetItemRate
                ndItem.setAttribute "hBinAndLocCheck",sCheckNoOfBinAndLoc
                ndItem.setAttribute "hLocNo",sLocNo
                ndItem.setAttribute "hBinNo",sBinNo
                ndItem.setAttribute "hMarketPrice",nGetMarketPrice
                ndItem.setAttribute "Counter",iCounter
                ndItem.setAttribute "ItemValue",dcrs(10)
                ndItem.setAttribute "FSNValue",sFSNFlag
                ndItem.setAttribute "ParCount",nNoOfSupplier
                ndItem.setAttribute "OrgPartyCode",sOrgnPartyCode
                ndItem.setAttribute "PartyName",sPartyName 
                ndItem.setAttribute "SuppParCode",sSuppParCode
                ndItem.setAttribute "SuppParType",sSuppParType
                ndItem.setAttribute "SuppParSubType",sSuppParSubType
                ndRoot.appendChild ndItem    
                iSNo = iSNo + 1
            dcrs.MoveNext
        loop
    end if
  '  Response.Clear 
    Response.ContentType = "text/xml"
    Response.Write oDOM.xml
%>