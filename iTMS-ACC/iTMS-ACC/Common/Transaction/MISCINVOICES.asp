<%@ Language="VBScript" %>
<% option explicit %>
<%
	'Program Name				:	ADVANCEPAYMNTREQUEST.asp
	'Module Name				:	PURCHASE
	'Author Name				:   RAGAVENDRAN R
	'Created On					:   April 07,2011
	'Modified By				:
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include File="../../include/Purpopulate.asp" -->
<!--#include File="../../include/sessionVerify.asp" -->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
    Dim rsObj,rsTemp
    Dim iRecCtr,iPrevPage,iTotalPages,nPageCtr,iNextPage,iCreateTransNo
    Dim iMiscNo,dVouDate,nVouAmt
    Dim sOrgID,sQuery,iPartyCode,sPartyName,sParType,sAppCode,sAppRefType,sAppRefNo,sAppRefDate
    Dim sFromDate,sToDate,sArrPeriod,sFinPeriod,sRefName,sRefNoDate,sMiscNoDate,sCreatedTransNo
    sOrgID = Session("organizationcode")
    sFinPeriod = Session("FinPeriod")

    set rsObj = Server.CreateObject("ADODB.Recordset")
    set rsTemp = Server.CreateObject("ADODB.Recordset")

        iPartyCode = Request.QueryString("ParCode")
        sFromDate = Request.QueryString("FromDate")
        sToDate = Request.QueryString("ToDate")
        sAppCode = Request.QueryString("APPCODE")

        if trim(sAppCode) = "" then
			sAppCode = Request.Form("hAppCode")
        end if

        if sAppCode=2 then
            sParType = "CR"
        else
            sParType = "DR"
        end if


        if Trim(sFromDate)="" then
            sArrPeriod = Split(sFinPeriod,":")
            sFromDate = "01/04/"& sArrPeriod(0)
            sToDate = "31/03/"& sArrPeriod(1)
        end if

    %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" ID="OutData">
	<Root/>
</script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/printwindow.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script>
function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function field(name) {
	var lower = String(name).toLowerCase();
	var form = document.formname;
	if (form.elements[name]) {
		return form.elements[name];
	}
	for (var i = 0; i < form.elements.length; i += 1) {
		if (String(form.elements[i].name).toLowerCase() === lower) {
			return form.elements[i];
		}
	}
	return null;
}

function dateControl(name) {
	return field(name) || document.getElementById(name);
}

function xmlRoot(value) {
	if (!value) {
		return null;
	}
	if (typeof value === "string") {
		return new DOMParser().parseFromString(value, "text/xml").documentElement;
	}
	if (value.nodeType === 1) {
		return value;
	}
	if (value.documentElement) {
		return value.documentElement;
	}
	if (value.XMLDocument && value.XMLDocument.documentElement) {
		return value.XMLDocument.documentElement;
	}
	if (value._doc && value._doc.documentElement) {
		return value._doc.documentElement;
	}
	return null;
}

function childElements(node, name) {
	var nodes = [];
	var wanted = name ? String(name).toLowerCase() : "";
	for (var i = 0; node && i < node.childNodes.length; i += 1) {
		if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
			nodes.push(node.childNodes[i]);
		}
	}
	return nodes;
}

function openModernDialog(url, args, features, callback) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	} else {
		window.open(url, "_blank", "height=500,width=600,resizable=yes,status=no");
	}
}

function runSupplierDialog(query, done) {
	openModernDialog("../SupplierSelect.asp?" + query, window.OutData || document.getElementById("OutData"), "status:no", function (outValue) {
		var root = xmlRoot(outValue);
		var action = trim(root && root.getAttribute("Action")).toUpperCase();
		var passQuery = trim(root && root.getAttribute("PassQuery"));
		if (!root || action === "CLOSE") {
			return;
		}
		if (action !== "DONE" && passQuery !== "") {
			runSupplierDialog(passQuery, done);
			return;
		}
		done(root);
	});
}

function selectedInvoice() {
	var count = parseInt(document.formname.hCnt.value, 10) || 0;
	var selected = [];
	for (var i = 1; i <= count; i += 1) {
		var item = field("ChkZ" + i);
		if (item && item.checked) {
			selected.push(String(item.value || "").split(":"));
		}
	}
	return selected;
}

function Create() {
	document.formname.action = "NewMiscInvoice.asp?APPCODE=" + encodeURIComponent(document.formname.hAppCode.value);
	document.formname.submit();
}

function EditInv() {
	var selected = selectedInvoice();
	var invNo;
	var crVouNo;
	if (selected.length !== 1) {
		alert("Select any one Invoice to Edit");
		return;
	}
	invNo = selected[0][0] || "";
	crVouNo = selected[0][1] || "";
	if (trim(crVouNo) !== "" && trim(crVouNo) !== "0") {
		alert("Voucher is Created Against the Misc. Invoice, So you can not edit this.");
		return;
	}
	document.formname.action = "MiscInvEdit.asp?APPCODE=" + encodeURIComponent(document.formname.hAppCode.value) + "&InvNo=" + encodeURIComponent(invNo);
	document.formname.submit();
}

function DeleteInv() {
	var selected = selectedInvoice();
	var invNo;
	var crVouNo;
	if (selected.length !== 1) {
		alert("Select any one Invoice to delete");
		return;
	}
	invNo = selected[0][0] || "";
	crVouNo = selected[0][1] || "";
	if (trim(crVouNo) !== "" && trim(crVouNo) !== "0") {
		alert("Voucher is Created Against the Misc. Invoice, So you can not delete this.");
		return;
	}
	document.formname.action = "MiscInvoiceDelete.asp?InvNo=" + encodeURIComponent(invNo);
	document.formname.submit();
}

function Supplierselect() {
	var query = "Unit=" + encodeURIComponent(document.formname.hOrgID.value) + "&hSelectMode=S&Flag=1&ParType=" + encodeURIComponent(document.formname.hParType.value);
	runSupplierDialog(query, function (root) {
		var suppliers = childElements(root, "Supplier");
		var codes = [];
		var names = [];
		for (var i = 0; i < suppliers.length; i += 1) {
			codes.push(trim(suppliers[i].getAttribute("SuppCode")));
			names.push(trim(suppliers[i].getAttribute("SuppName")));
		}
		document.getElementById("idSupplier").innerHTML = names.join(",");
		document.formname.hSupplierCode.value = codes.join(",");
		document.formname.hSupplierName.value = names.join(",");
	});
}

function Validate() {
	var parCode = document.formname.hSupplierCode.value;
	var fromDate = dateControl("ctlFromDate").getDate();
	var toDate = dateControl("ctlToDate").getDate();
	document.formname.action = "MISCINVOICES.ASP?ParCode=" + encodeURIComponent(parCode) + "&FromDate=" + encodeURIComponent(fromDate) + "&ToDate=" + encodeURIComponent(toDate);
	document.formname.submit();
}

function parseDate(value) {
	var parts = String(value || "").split("/");
	if (parts.length === 3) {
		return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
	}
	return null;
}

function Setdate() {
	var fromDate = document.formname.hFromDate.value;
	var toDate = document.formname.hToDate.value;
	var toDateObj = parseDate(toDate);
	var today = new Date();
	var fromControl = dateControl("ctlFromDate");
	var toControl = dateControl("ctlToDate");
	if (toDateObj && toDateObj < new Date(today.getFullYear(), today.getMonth(), today.getDate())) {
		fromControl.setDate(fromDate);
		toControl.setDate(toDate);
	} else {
		fromControl.setMinDate(fromDate);
		toControl.setMaxDate(today);
	}
}

function setdate() {
	Setdate();
}

function AssignPage(nPage) {
	document.formname.hPage.value = nPage;
	document.formname.submit();
}

function ResetData() {
	document.formname.hSupplierCode.value = "";
	document.formname.hSupplierName.value = "";
	document.getElementById("idSupplier").innerHTML = "";
	Setdate();
}
</script>
</head>

<body leftmargin="0" topmargin="0"  onload="Setdate()">
	<form method="POST" name="formname" action="">
	<input type=hidden name="hPage" value="">
	<input type=hidden name="hOrgID" value="<%=sOrgID%>">
	<input type=hidden name="hSupplierName" value="<%=sPartyName%>">
	<input type=hidden name="hSupplierCode" value="<%=iPartyCode%>">
	<input type=hidden name="hFromDate" value="<%=sFromDate%>">
	<input type=hidden name="hToDate" value="<%=sToDate%>">
	<input type=hidden name="hParType" value="<%=sParType%>">
	<input type=hidden name="hAppCode" value="<%=sAppCode%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Misc. Invoices
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>


<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="center" class="SubTitle">&nbsp;&nbsp;
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
<tr>
<td width="100%">
<div id="idUnprocessed" style="width: 575; display: none">
<table cellpadding="0" cellspacing="0">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="7">
</td>
</tr>

<!--<tr>
	<td class="FieldCellSub">	</td>
	<td class="FieldCellSub">    Item Type	</td>
	<td class="FieldCellSub" colspan="2">
    <select size="1" name="selItemType" class="FormElem">
		<%'popSelItemType(sItemType)	%>
	</select>
	</td>
</tr>-->

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
		Date From
	</td>
	<td class="FieldCellSub">

	<% Response.Write InsertDatePicker("ctlFromDate") %>

	</td>
	<td class="FieldCellSub" colspan="2"></td>
	<td class="FieldCellSub" colspan="2">To	</td>
	<td class="FieldCellSub">

	<%Response.Write InsertDatePicker("ctlToDate") %>

	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<% if sAppCode=2 then
	        Response.Write "Supplier"
	   else
	        Response.Write "Party"
	   end if
	%>


	</td>
	<td class="FieldCellSub">
	<span class="dataonly" id="idSupplier"></span> &nbsp;
	&nbsp;<img id="Img1" border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" onclick="Supplierselect()" align="middle" alt="Supplier Selection" width="10" height="11">
	</td>

</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">

        &nbsp;

	</td>
	<td class="FieldCellSub">

        <p align="right">
            <input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
        </p>
    </td>
	<td class="FieldCellSub" colspan="2"></td>

	<td class="FieldCellSub" colspan="3">
	<input type="button" value="Reset" name="Cmdreset" class="ActionButtonX" onclick="ResetData()" >
	</td>
</tr>

</table>
</div>
</td>
</table>
</div>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top">
<table border="0" cellspacing="1" class="ExcelTable" width="100%">
<tr>
<td class="ExcelHeaderCell" align="center" rowspan=2 >S.No.</td>
<td class="ExcelHeaderCell" align="center" rowspan=2 >
    <img src="../../assets/images/iTMS%20icons/Deleteicon.gif" alt="Click here to delete the Advance Payment Request" onclick="DeleteInv()">
</td>
<%if sAppCode = 2 then %>
    <td class="ExcelHeaderCell" align="center" rowspan=2 >Supplier Name</td>
<%else %>
    <td class="ExcelHeaderCell" align="center" rowspan=2 >Party Name</td>
<%end if %>
<td class="ExcelHeaderCell" align="center" colspan=2>Reference</td>
<td class="ExcelHeaderCell" align="center" colspan=2>Misc. Invoice</td>
</tr>
<tr>
    <td class="ExcelHeaderCell" align="center">Type</td>
    <td class="ExcelHeaderCell" align="center">No - Date</td>
    <td class="ExcelHeaderCell" align="center">No - Date</td>
    <td class="ExcelHeaderCell" align="center">Amount</td>
</tr>
<%
	'Response.Write "<font color=red>"
        sQuery = "Select MiscTransNo,Convert(varchar,VoucherDate,103),VoucherAmount,PartyCode,"&_
        " AppRefType,AppRefNo,AppRefDate,isNull(CreatedMiscPymtNo,MiscTransNo),IsNull(ReceiptNo,0) as CreatedTransNo from Acc_T_MiscPymtRequestHeader where ApplicationCode = "& sAppCode

        if trim(iPartyCode)<>"" then
           sQuery = sQuery & " and PartyCode = "& iPartyCode
        end if
		'Response.Write "<p> " & sQuery
        rsObj.Open sQuery,con
        if not rsObj.EOF then
            do while not rsObj.EOF
                iRecCtr = iRecCtr + 1
                    iMiscNo = rsObj(0)
                    dVouDate =rsObj(1)
                    nVouAmt = rsObj(2)
                    iPartyCode = rsObj(3)
                    sAppRefType = rsObj(4)
                    sAppRefNo = rsObj(5)
                    sAppRefDate = rsObj(6)
                    sRefNoDate = rsObj(7)
                    sCreatedTransNo = rsObj(8)
                    sMiscNoDate = iMiscNo &" - "& dVouDate

                        sQuery = "Select PartyName from VwOrgParty where PartyCode = "& iPartyCode
                        rsTemp.Open sQuery,con
                        if not rsTemp.EOF then
                            sPartyName = trim(rsTemp(0))
                        end if
                        rsTemp.Close
                        if Trim(sAppRefType)<>"" then
							sQuery = "Select ReferenceName from VW_ReferenceTypes where ReferenceEntryNo ="& sAppRefType
							'Response.Write sQuery
							rsTemp.Open sQuery,con
							if not rsTemp.EOF then
							    sRefName = trim(rsTemp(0))
							end if
							rsTemp.Close
						end if'if Trim(sAppRefType)<>"" then

                    %>
                        <tr>
                            <td class="ExcelSerial" align="center"><%=iRecCtr%></td>
                            <td class="ExcelDisplayCell" align="center">
                                <input type=checkbox name="chkZ<%=iRecCtr%>" value="<%=iMiscNo%>:<%=sCreatedTransNo%>">
                            </td>
                            <td class="ExcelDisplayCell" align="left"><%=sPartyName%></td>
                            <td class="ExcelDisplayCell" align="left"><%=sRefName%></td>
                            <td class="ExcelDisplayCell" align="left"><%=sRefNoDate%></td>
                            <td class="ExcelDisplayCell" align="left"><%=sMiscNoDate%></td>
                            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(nVouAmt,2)%></td>
                        </tr>
                    <%
                rsObj.MoveNext
            loop
        end if
        rsObj.Close
%>
</table>
</td>
<td align="center" class="ClearPixel" width="5">
<input type="hidden" name="hCnt" value="<%=iRecCtr%>" >
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top" align="right">

<input type="button" value=" |< " class="ActionButtonX" id=ButFirst name=ButFirst onClick="AssignPage('1')">

<%if trim(iPrevPage) = "0" then  %>
	<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev >
<%else%>
	<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev onClick="AssignPage('<%=iPrevPage%>')">
<%end if %>


<SELECT class="FormElem" onChange="AssignPage(this.value)"  id="mCmbPage" name="mCmbPage">

<%for nPageCtr= 1 to iTotalPages %>
	<option value="<%=nPageCtr%>" <%if trim(iPageNo) = trim(nPageCtr) then Response.Write "Selected" %> >Page <%=nPageCtr%> of <%=iTotalPages %></option>
<%next%>

</SELECT>
<%if trim(iNextPage) = "0" then  %>
	<input type="button" value=" >> " class="ActionButtonX" id=ButNext name=ButNext >
<%else%>
	<input type="button" value=" >> " class="ActionButtonX" onclick="AssignPage('<%=iNextPage%>')" id=ButNext name=ButNext >
<%end if%>

<input type="button" value=" >| " class="ActionButtonX" id=ButLast name=ButLast OnClick="AssignPage('<%=iTotalPages %>')">

</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td valign="middle" class="ActionCell">
    <input type="button" value="New Invoice" name="btnRequest" class="ActionButtonX" onclick="Create()">
    <input type="button" value=" Edit " name="btnEdit" class="ActionButtonX" onclick="EditInv()">
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="BottomPack" colspan="3">
</td>
</tr>

</table>
</td>
</tr>

</table>
</td>
</tr>

</table>
</form>
</body>
</html>


<%
Function SelFun()
Dim sTemp
	sTemp = Request.Form("hChoice")
	'Response.Write sTemp
	If sTemp = "A" or sTemp = "C" then
		Response.Write ("<OPTION VALUE=""CRW"">Create (w/t Ref. based)</OPTION>" &vbcrlf)
		Response.Write ("<OPTION VALUE=""EDT"">Edit</OPTION>"  &vbcrlf)
		Response.Write ("<OPTION VALUE=""DEL"">Delete</OPTION>" &vbcrlf)
		'Response.Write ("<OPTION VALUE=""APP"">Finalise</OPTION>" &vbcrlf)
	ElseIf sTemp = "D" then
		Response.Write ("<OPTION VALUE=""CRW"">Create (w/t Ref. based)</OPTION>" &vbcrlf)
		Response.Write ("<OPTION VALUE=""EDT"">Edit</OPTION>"  &vbcrlf)
		Response.Write ("<OPTION VALUE=""DEL"">Delete</OPTION>" &vbcrlf)
		'Response.Write ("<OPTION VALUE=""APP"">Finalise</OPTION>" &vbcrlf)
	ElseIf sTemp = "P" then
		Response.Write ("<OPTION VALUE=""CRE"">Create (Ref. based)</OPTION>" &vbcrlf)
		Response.Write ("<OPTION VALUE=""CRW"">Create (w/t Ref. based)</OPTION>" &vbcrlf)
	Else

		Response.Write ("<OPTION VALUE=""CRE"">Create (Ref. based)</OPTION>" &vbcrlf)
		Response.Write ("<OPTION VALUE=""CRW"">Create (w/t Ref. based)</OPTION>" &vbcrlf)
	End IF

End Function
%>
