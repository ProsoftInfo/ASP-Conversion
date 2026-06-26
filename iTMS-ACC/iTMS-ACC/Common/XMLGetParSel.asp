<%@language="VBScript"%>
<%Option Explicit%>
<%
	'Program Name				:	XMLGetParSel.asp
	'Module Name				:	To Diplay Employee
	'Author Name				:	Ragavendran R
	'Created On					:	March 22,2013
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
Dim sOrgID,sFilter,sSearchBy,sSelectMode,sFlag,sQuery
Dim sFinPeriod,sFinYearFrom,sFinYearTo,sTempMonYr,sMonYr,sPartyCode
Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,iSNo,iPageSize
Dim dcrs,rsTemp,rsTemp1,sTemp,sRequest,iCounter
Dim sSearchEmpCode,sSearchEmpName

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
sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
'Response.Write "<p><font color=red>"

sSearchEmpCode = Request("EmpCode")
sSearchEmpName  = Request("EmpName")

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

sQuery = "Select OrgnPartyCode,PartyCode,PartyName,PartyName from APP_M_PartyMaster where 1 = 1 "

if trim(sSearchEmpCode)<>"" then
    sQuery = sQuery &" and OrgnPartyCode like '%"& sSearchEmpCode&"%'"
end if

if trim(sSearchEmpName)<>"" then
    sQuery = sQuery &" and PartyName like '%"& sSearchEmpName &"%'"
end if


sQuery = sQuery & " Group by OrgnPartyCode,PartyCode,PartyName"

' Response.write sQuery

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
            set ndItem = oDOM.createElement("Emp")
                ndItem.setAttribute "SNo",iSNo
                ndItem.setAttribute "Counter",iCounter
                ndItem.setAttribute "EmpCode",dcrs(0)
                ndItem.setAttribute "EmpID",dcrs(1)
                ndItem.setAttribute "EmpFullName",dcrs(2)
                ndItem.setAttribute "EmpName",dcrs(3)
                ndRoot.appendChild ndItem    
                iSNo = iSNo + 1
            dcrs.MoveNext
        loop
    end if
    
    Response.ContentType = "text/xml"
    Response.Write oDOM.xml
%>