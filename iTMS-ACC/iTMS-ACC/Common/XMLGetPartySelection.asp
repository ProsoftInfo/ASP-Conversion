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

<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/CommonFunctions.asp"-->
<%

Dim sTable,oDOM
Dim sOrgID,sFilter,sSearchBy,sSelectMode,sFlag,sQuery
Dim sFinPeriod,sFinYearFrom,sFinYearTo,sTempMonYr,sMonYr,sPartyCode
Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,iSNo,iPageSize
Dim dcrs,rsTemp,rsTemp1,sTemp,sRequest,iCounter
Dim sSearchOrgParCode,sSearchPartyName,sSearchParType,sArrParType

Dim ndRoot,ndItem
Dim sOrgnPartyCode,sPartyName
Dim sParType,sParSubType,sParTypeName

Set dcrs = Server.CreateObject("ADODB.Recordset")
Set rsTemp = Server.CreateObject("ADODB.Recordset")
Set rsTemp1 = Server.CreateObject("ADODB.Recordset")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

set ndRoot = oDOM.createElement("Root")
oDOM.appendChild ndRoot

'Response.Write "<font color=red>"

sOrgID = trim(Request.QueryString("orgID"))
sTemp=split(trim(Request("Party")),"?")
sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
'Response.Write "<p><font color=red>"

if trim(Request("Party"))<>"" then
    if UBound(sTemp)>1 then
	    sParType=sTemp(0)
        sParSubType=sTemp(1)
        sParTypeName=sTemp(2)
    else
        sParType = sTemp(0)
    end if
end if 'if trim(Request("Party"))<>"" then


sSearchOrgParCode = Request("OrgnPartyCode")
sSearchPartyName  = Request("PartyName")
sSearchParType    = Request("ParType")

if trim(sSearchParType)="0|0" then sSearchParType = ""

if trim(sSearchParType)<>"" then
    sArrParType = split(sSearchParType,"|")
end if

sFilter = trim(Request.QueryString("Query"))&"%"

iPageSize = Request.QueryString("PageSize")
if Trim(iPageSize)="" or IsNull(iPageSize) then iPageSize = 15

iCurrentPage=Request("Page")
if Trim(iCurrentPage)="" or IsNull(iCurrentPage) then iCurrentPage = 1
iCurrentPage = CInt(iCurrentPage)


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

iSAApplicationPop = Session("iApplication")
iSAProcessPop = Session("iProcess")
iSAActivityPop = Session("iActivity")
iEmpNoPopulate = Session("employeenumber")

sQuery = "Select OrgnPartyCode,PartyCode,PartyName from VwOrgParty where OUDefinitionID='"&sOrgId&"'"
if trim(sParType)<>"" then
    sQuery = sQuery &" and PartyType = '"& sParType &"'"
end if

if trim(sSearchOrgParCode)<>"" then
    sQuery = sQuery &" and OrgnPartyCode like '%"& sSearchOrgParCode&"%'"
end if

if trim(sSearchPartyName)<>"" then
    sQuery = sQuery &" and PartyName like '%"& sSearchPartyName &"%'"
end if

if trim(sSearchParType)<>"" then
    sQuery = sQuery &" and PartyType = '"& sArrParType(0) &"' and PartySubType = "& sArrParType(1)
end if

sQuery = sQuery & " Group by OrgnPartyCode,PartyCode,PartyName"

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
            set ndItem = oDOM.createElement("Party")
                ndItem.setAttribute "SNO",iSNo
                ndItem.setAttribute "Counter",iCounter
                ndItem.setAttribute "OrgnPartyCode",dcrs(0)
                ndItem.setAttribute "PartyCode",dcrs(1)
                ndItem.setAttribute "PartyName",dcrs(2)
                ndRoot.appendChild ndItem    
                iSNo = iSNo + 1
            dcrs.MoveNext
        loop
    end if
    
    Response.ContentType = "text/xml"
    Response.Write oDOM.xml
%>