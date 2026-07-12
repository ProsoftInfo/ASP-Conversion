<%@language="VBScript"%>
<%Option Explicit%>
<%
	'Program Name				:	ItemSelectRelPartyCommon.asp
	'Module Name				:	To List all Item details and Display SuppItem Code,Desc for the Related item with Selected Party
	'Author Name				:	Ragavendran R
	'Created On					:	Dec 09,2011
	'Modified By				:	Ragavendran R
	'Modified On				:	Aug 23,2012 'Additionaly adding the BOM Sub level to the Item Node 
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

Dim sTable
Dim sIType,sOrgID,sFilter,sSearchBy,sSelectMode,sFlag,sQuery,sFlagItemStock
Dim sFinPeriod,sFinYearFrom,sFinYearTo,sTempMonYr,sMonYr,sButtDispMode,sPartyCode
Dim sSuppItemCode,sSuppItemDesc
Dim iStock,iClassCodes,iCounter
Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,iSNo
Dim dcrs,rsTemp
Dim DisableBut,sCheckNoOfBinAndLoc,sFinYrFrom,sFinYrTo,sLocNo,sBinNo
Dim nGetItemRate,nGetMarketPrice,sRequest,sSearchType,sPartyType,sCallFrom
Dim sDisplayItem


Set dcrs = Server.CreateObject("ADODB.Recordset")
Set rsTemp = Server.CreateObject("ADODB.Recordset")

Const iPageSize = 12
Response.Write "<font color=red>"

sOrgID = trim(Request.QueryString("orgID"))
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
sPartyType = Request.QueryString("PartyType")
sCallFrom = Request.QueryString("CallFrom")
sDisplayItem = Request.QueryString("Disp")

if Trim(sSearchType)="" or IsNull(sSearchType) then sSearchType = "C"
if Trim(sSearchBy)="" or IsNull(sSearchBy) then sSearchBy = "IC"

sRequest = "orgID="&sOrgID&"&sIType="&sIType&"&hSelectMode="&sSelectMode&"&Flag="&sFlag&"&hClassCodes="&iClassCodes&"&hDispButt="&sButtDispMode&"&hDispItem="& sFlagItemStock &"&hPartyCode="&sPartyCode&"&PartyType="&sPartyType &"&CallFrom="&sCallFrom&"&Disp="&sDisplayItem


iCurrentPage=Request("Page")
if Trim(iCurrentPage)="" or IsNull(iCurrentPage) then iCurrentPage = 1
iCurrentPage = CInt(iCurrentPage)

DisableBut = "N"
if Trim(sSelectMode)="S" or Trim(sSelectMode)="R" then
    DisableBut = "N"
else
    DisableBut = "Y"
end if

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
'sButtDispMode = "Y"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Selection</TITLE>
<base target="_self"></base>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="TempItem"><Root CurrPage="1" TotPage="1"></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="XMLAttributeList"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="BOMItem"><Root/></script>
<script src="../scripts/itms-modern-compat.js"></script>
<script src="../scripts/rolloverout.js"></script>
<script src="../scripts/DataValidation.js"></script>
<script>
function DoKeyPress(evt, sYesNo, iIntPart, iDecPart) {
	evt = evt || {};
	var target = evt.target || null;
	var key = evt.keyCode || evt.which || 0;
	var value = target ? String(target.value || "") : "";
	var decimalPosition = value.indexOf(".");
	var intValue = decimalPosition >= 0 ? value.substring(0, decimalPosition) : value;
	var decValue = decimalPosition >= 0 ? value.substring(decimalPosition + 1) : "";
	var isNumeric = key >= 48 && key <= 57;
	var isCompare = key === 60 || key === 61 || key === 62;
	var block = false;

	if (sYesNo === "N") {
		block = (!isNumeric && !isCompare) || intValue.length >= iIntPart;
	} else if (sYesNo === "Y") {
		block = (!isNumeric && key !== 46 && !isCompare) ||
			(decimalPosition >= 0 && (key === 46 || decValue.length >= iDecPart)) ||
			(decimalPosition < 0 && intValue.length >= iIntPart && key !== 46);
	}
	if (block) {
		if (evt.preventDefault) {
			evt.preventDefault();
		}
		return false;
	}
	return true;
}
</script>
<script src="../scripts/ItemSelectRelPartyCommonCompat.js"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 onload="Init()" onkeydown="CallSearchMain(event)">
<form method="POST" name="formname" class="PopupTable">
<Input type="hidden" name="hOrgID" value="<%=sOrgID%>">
<input type="hidden" name="hTemp" value="<%=Request.QueryString%>">
<input type="hidden" name="hSelectMode" value="<%=sSelectMode%>">
<input type="hidden" name="hDisableBut" value="<%=DisableBut%>">
<input type="hidden" name="hPartyCode" value="<%=sPartyCode%>">
<input type="hidden" name="hPage" value="0">
<input type="hidden" name="hRequest" value="<%=sRequest%>">
<input type="hidden" name="hButtDispMode" value="<%=sButtDispMode%>">
<input type="hidden" name="hChkCount" value="0">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">
<table border="0" width="98%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
	    <td width="10"></td>
		<td valign="top">
			<table id="tblItem" border="0" cellpadding="0" cellspacing="1" class="ExcelTable" width="100%">
			    <tr>
			        <td class="ExcelHeaderCell" align="center">Select</td>
			        <td class="ExcelHeaderCell" align="center">Item Code</td>
			        <td class="ExcelHeaderCell" align="center">Item Description</td>
			        <td class="ExcelHeaderCell" align="center">Party Item Code</td>
			        <td class="ExcelHeaderCell" align="center">Party Item Description</td>
			        <td class="ExcelHeaderCell" align="center">Stock</td>
			        <td class="ExcelHeaderCell" align="center">Group Name</td>
			        <td class="ExcelHeaderCell" align="center">UOM</td>
			        <td class="ExcelHeaderCell" align="center">Attribute List</td>
			    </tr>
			    <tr>
					<td class="ExcelHeaderCell" align="center"></td>
					<td class="ExcelHeaderCell" align="left"><input type="text" name="txtSearchItemCode" class="FormElem" size="15" onblur="ShowPage('<%=sRequest%>&Page='+document.formname.hPage.value)" onkeyup="CallSearch()" tabindex="1"></td>
					<td class="ExcelHeaderCell" align="left">
						<input type="checkbox" name="chkExact"  tabindex="2">Start with <input type="text" name="txtSearchItemName" class="FormElem" size="24" onblur="ShowPage('<%=sRequest%>&Page='+document.formname.hPage.value)" onkeyup="CallSearch()"  tabindex="3">
					</td>
					<td class="ExcelHeaderCell" align="left"><input type="text" name="txtSearchPartyItemCode" class="FormElem" size="15" onblur="ShowPage('<%=sRequest%>&Page='+document.formname.hPage.value)" onkeyup="CallSearch()"   tabindex="4"></td>
					<td class="ExcelHeaderCell" align="left"><input type="text" name="txtSearchPartyItemName" class="FormElem" onblur="ShowPage('<%=sRequest%>&Page='+document.formname.hPage.value)" onkeyup="CallSearch()"   tabindex="5"></td>
					<td class="ExcelHeaderCell" align="left">
					    <input type="text" name="txtSearchStock" class="FormElem" onblur="CallSearchStock()" onkeypress="return DoKeyPress(event,'Y',7,2);" onkeyup="CallSearchStock(event)" tabindex="6" size="8" value="">
					</td>
					<td class="ExcelHeaderCell" align="left"></td>
					<td class="ExcelHeaderCell" align="center"></td>
					<td class="ExcelHeaderCell" align="center"></td>
			    </tr>
			    
			    <tr>
                    <td valign="top" class="ExcelHeaderCell" align="center" colspan="9">Page&nbsp;
                    <input type=text class="FormElem" size=5 style="text-align:right" name="txtCurrPage" onblur="ShowPage('<%=sRequest%>&Page='+this.value)" onkeydown="CallSearch(event)" >&nbsp;
                    of&nbsp;<span id="spanTotPage"></span>
                    </td>
                </tr>
                <tr>
                    <td class="ExcelHeaderCell" colspan="9" align="center">
                        <%if sButtDispMode="Y" then %>
                            <input type="button" name="btnAddNew" value="Add New" onclick="WithOutMat()" class="ActionButtonX">
                        <%end if %>
                        <%if sSelectMode ="M" then %>
                            <input type="button" name="btnAddToList" value="Add To List" class="ActionButtonX" onclick="AddFun()" disabled >
                            <input type="button" name="btnDone" value="Done" class="ActionButtonX" onclick="SendValue()">
                        <%else %>
                            <input type="button" name="btnAddToList" value="Add To List" disabled class="ActionButtonX">
                            <input type="button" name="btnDone" value="Done" class="ActionButtonX" onclick="SendValue()">
                        <%end if%>
                    </td>
                </tr>
                <tr>
                    <td class="ExcelHeaderCell" colspan="9">
                     Selected Entries<span id="idSelList"></span>
                    </td>
                </tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</html>
<%
Private Function PageQuery(ByRef pnPage)
		Dim lsQuery,iCount

		'lsQuery = Request.QueryString
		lsQuery = sRequest

		iCount = InStr(1,lsQuery,"&Query=")
		If cint(iCount) > 0 Then
			lsQuery = left(lsQuery,iCount - 1)
		End If

		lsQuery = Replace(lsQuery, "Page=" & lnPage, "")
		'Response.Write lsQuery

		If pnPage < 1 Then
			pnPage = 1
		ElseIf pnPage > iTotalPage Then
			pnPage = iTotalPage
		End If

		If lsQuery = "" Then
			lsQuery = "Page=" & pnPage
		ElseIf Right(lsQuery, 1) = "&" Then
			lsQuery = lsQuery & "Page=" & pnPage
	   	Else
			lsQuery = lsQuery & "&Page=" & pnPage
		End If

		'PageQuery = "?" & lsQuery
		PageQuery = lsQuery
		'Response.Write PageQuery

	End Function
%>
