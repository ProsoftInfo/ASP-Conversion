<%@language="VBScript"%>
<%Option Explicit%>
<%
	'Program Name				:	ParSelPop.asp
	'Module Name				:	Admin(Master)
	'Author Name				:	Ragavendran R
	'Created On					:	March 22,2013
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

Dim sTable
Dim sIType,sOrgID,sFilter,sSearchBy,sSelectMode,sQuery
Dim sFinPeriod,sFinYearFrom,sFinYearTo,sTempMonYr,sMonYr,sEmpID
Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,iSNo
Dim dcrs,rsTemp,sTemp,sRequest,sEmpCode,sEmpName

Set dcrs = Server.CreateObject("ADODB.Recordset")
Set rsTemp = Server.CreateObject("ADODB.Recordset")

Const iPageSize = 15
Response.Write "<font color=red>"

sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
Response.Write "<p><font color=red>"

sFilter = trim(Request.QueryString("Query"))&"%"

if Trim(sSearchBy)="" or IsNull(sSearchBy) then sSearchBy = "IC"

sRequest = "SearchBy="&sSearchBy&"&Query="&Request("Query")&"&hSelectMode="&sSelectMode

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

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Employee Selection</TITLE>
<base target="_self"></base>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="TempItem"><Root CurrPage="1" TotPage="1"></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="XMLPartySubType"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>

<script src="/Scripts/itms-modern-compat.js"></script>
<script SRC="../scripts/rolloverout.js"></SCRIPT>
<script SRC="../scripts/DataValidation.js"></SCRIPT>
<script>
function itmsRememberPartyKeyEvent(evt) {
	window.__itmsPartyKeyEvent = evt || null;
}
if (document.addEventListener) {
	document.addEventListener("keypress", itmsRememberPartyKeyEvent, true);
	document.addEventListener("keydown", itmsRememberPartyKeyEvent, true);
}
function DoKeyPress(sYesNo, iIntPart, iDecPart, evt) {
	var eventObj = evt || window.__itmsPartyKeyEvent || null;
	var target = eventObj && eventObj.target || null;
	var code = keyCodeFromEvent(eventObj);
	var value = target && target.value != null ? String(target.value) : "";
	var decPosition = value.indexOf(".");
	var intValue = decPosition >= 0 ? value.substring(0, decPosition) : value;
	var decValue = decPosition >= 0 ? value.substring(decPosition + 1) : "";
	var isControl = code === 0 || code === 8 || code === 9 || code === 13 || code === 27;
	var isCompare = code === 60 || code === 61 || code === 62;
	function cancelKey() {
		if (eventObj && eventObj.preventDefault) {
			eventObj.preventDefault();
		}
		return false;
	}
	if (isControl) {
		return true;
	}
	if (sYesNo === "N" && ((code < 48 || code > 57) && !isCompare)) {
		return cancelKey();
	}
	if (sYesNo === "Y" && ((code < 48 || code > 57) && code !== 46 && !isCompare)) {
		return cancelKey();
	}
	if (sYesNo === "N" && intValue.length >= iIntPart) {
		return cancelKey();
	}
	if (sYesNo === "Y") {
		if (decPosition >= 0) {
			if (code === 46 || decValue.length >= iDecPart) {
				return cancelKey();
			}
		} else if (intValue.length >= iIntPart && code !== 46) {
			return cancelKey();
		}
	}
	return true;
}
function keyCodeFromEvent(eventObj) {
	if (!eventObj || !eventObj.key) {
		return 0;
	}
	if (eventObj.key.length === 1) {
		return eventObj.key.charCodeAt(0);
	}
	if (eventObj.key === "Backspace") {
		return 8;
	}
	if (eventObj.key === "Tab") {
		return 9;
	}
	if (eventObj.key === "Enter") {
		return 13;
	}
	if (eventObj.key === "Escape") {
		return 27;
	}
	return 0;
}
</Script>

<script>
window.__itmsPartySelectorConfig = {
	kind: "employee",
	dataUrl: "XMLGetParSel.asp"
};
</script>
<script src="../scripts/PartySelectorCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 onload="Init();" onkeydown="CallSearchMain(event)">
<form method="POST" name="formname" class="PopupTable">
<input type="hidden" name="hTemp" value="<%=Request.QueryString%>">
<input type="hidden" name="hSelectMode" value="<%=sSelectMode%>">
<input type="hidden" name="hEmpID" value="<%=sEmpID%>">
<input type="hidden" name="hPage" value="0">
<input type="hidden" name="hRequest" value="<%=sRequest%>">
<input type="hidden" name="hChkCount" value="0">

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
			        <td class="ExcelHeaderCell" align="center">Employee Code</td>
			        <td class="ExcelHeaderCell" align="center">Employee Name</td>
			    </tr>
			    <tr>
			        <td class="ExcelHeaderCell" align="center"></td>
			        <td class="ExcelHeaderCell" align="center"><input type="text" name="txtEmpCode" class="FormElem" onblur="ShowPage('<%=sRequest%>&Page='+document.formname.hPage.value)" onkeyup="CallSearch()"></td>
			        <td class="ExcelHeaderCell" align="center"><input type="text" name="txtEmpName" class="FormElem" onblur="ShowPage('<%=sRequest%>&Page='+document.formname.hPage.value)" onkeyup="CallSearch()"></td>
			    </tr>
			    <tr>
                    <td valign="top" class="ExcelHeaderCell" align="center" colspan="8">Page&nbsp;
                    <input type=text class="FormElem" size=5 style="text-align:right" name="txtCurrPage" onblur="ShowPage('<%=sRequest%>&Page='+this.value)" onkeydown="CallSearch()" >&nbsp;
                    of&nbsp;<span id="spanTotPage"></span>
                    </td>
                </tr>
                <tr>
                    <td class="ExcelHeaderCell" colspan="8" align="center">
                        <%if sSelectMode ="M" then %>
                            <input type="button" name="btnAddToList" value="Add To List" class="ActionButtonX" onclick="AddFun()">
                            <input type="button" name="btnDone" value="Done" class="ActionButtonX" onclick="SendValue()">
                        <%else %>
                            <input type="button" name="btnAddToList" value="Add To List" disabled class="ActionButtonX">
                            <input type="button" name="btnDone" value="Done" class="ActionButtonX" onclick="SendValue()">
                        <%end if%>
                    </td>
                </tr>
                <tr>
                    <td class="ExcelHeaderCell" colspan="8">
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
