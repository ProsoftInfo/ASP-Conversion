<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 27,2003
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
<%
	Dim sFinPeriod,sFromYr,sToYr,sTempYr
	sFinPeriod = Session("FinPeriod")
	IF CStr(sFinPeriod) <> "" Then
		sTempYr = Split(sFinPeriod,":")
		sFromYr = sTempYr(0)
		sToYr = sTempYr(1)
	End IF

%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML ID="OutData"><PartyType/></XML>
<XML id="AccHeadData">
<account/>
</XML>
<script src="../../scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/VouTransactions.js"></script>
<script>
var sFlag = "";

function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function formField(name) {
	var form = document.formname;
	return form && (form.elements[name] || form[name]) || null;
}

function attr(node, nameOrIndex) {
	var value;
	if (!node || !node.attributes) {
		return "";
	}
	if (typeof nameOrIndex === "number") {
		value = node.attributes.item(nameOrIndex);
	} else {
		value = node.attributes.getNamedItem(nameOrIndex);
	}
	return value ? value.value : "";
}

function childElements(node) {
	var nodes = [];
	for (var i = 0; node && i < node.childNodes.length; i += 1) {
		if (node.childNodes[i].nodeType === 1) {
			nodes.push(node.childNodes[i]);
		}
	}
	return nodes;
}

function islandRoot(name) {
	if (window.ITMSModernCompat) {
		window.ITMSModernCompat.upgradeXmlIslands(document);
	}
	var island = window[name] || document[name] || document.getElementById(name);
	return island && island.documentElement || island && island.XMLDocument && island.XMLDocument.documentElement || island && island._doc && island._doc.documentElement || null;
}

function xhrText(url) {
	var request = new XMLHttpRequest();
	request.open("GET", url, false);
	request.send(null);
	return request.responseText || "";
}

function loadIsland(name, text) {
	if (window.ITMSModernCompat) {
		window.ITMSModernCompat.upgradeXmlIslands(document);
	}
	var island = window[name] || document[name];
	if (island && typeof island.loadXML === "function") {
		island.loadXML(text);
	}
}

function controlDate(name) {
	var control = formField(name) || document.getElementById(name);
	if (control && typeof control.GetDate === "function") {
		return control.GetDate();
	}
	if (control && typeof control.getDate === "function") {
		return control.getDate();
	}
	return control ? control.value : "";
}

function parseLegacyDate(value) {
	var text = trim(value);
	var match;
	if (!text) {
		return null;
	}
	match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
	if (match) {
		return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
	}
	match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
	if (match) {
		var year = Number(match[3]);
		if (year < 100) {
			year += 2000;
		}
		return new Date(year, Number(match[2]) - 1, Number(match[1]));
	}
	return null;
}

function selectText(select) {
	return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
}

function setSpan(id, value) {
	var span = document.getElementById(id);
	if (span) {
		span.innerHTML = value || "";
	}
}

function isNumeric(value) {
	return trim(value) !== "" && !isNaN(Number(String(value).replace(/,/g, "")));
}

function numericValue(value) {
	return Number(String(value == null ? "" : value).replace(/,/g, ""));
}

function popPartType() {
	var select = formField("SelAccHead");
	if (!select) {
		return;
	}
	select.options.length = 2;
	if (document.formname.selUnitId.value !== "0") {
		var iUnitNo = document.formname.selUnitId.value;
		var responseText = xhrText("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(iUnitNo));
		if (trim(responseText) !== "") {
			loadIsland("OutData", responseText);
			childElements(islandRoot("OutData")).forEach(function (headerNode) {
				select.options[select.options.length] = new Option(headerNode.textContent || headerNode.text || "", attr(headerNode, "ParType"));
			});
		}
	}
}

function DisplayBook() {
	var bookSelect = formField("selBook");
	bookSelect.options.length = 1;
	if (document.formname.selUnitId.selectedIndex !== 0 && document.formname.selVoucher.selectedIndex !== 0) {
		var iUnitNo = document.formname.selUnitId.value;
		var bkCode = document.formname.selVoucher.value;
		var responseText = xhrText("XMLGetOrgBook.asp?BkCode=" + encodeURIComponent(bkCode) + "&orgID=" + encodeURIComponent(iUnitNo));
		if (trim(responseText) !== "") {
			loadIsland("UnitBookData", responseText);
			childElements(islandRoot("UnitBookData")).forEach(function (headerNode) {
				bookSelect.options[bookSelect.options.length] = new Option(attr(headerNode, 1), attr(headerNode, 0));
			});
		}
	}
}

function validate() {
	if (document.formname.selUnitId.selectedIndex < 1) {
		alert("Select Unit");
		document.formname.selUnitId.focus();
		return;
	}
	if (document.formname.selVoucher.selectedIndex < 1) {
		alert("Select Voucher type");
		document.formname.selVoucher.focus();
		return;
	}
	if (document.formname.selBook.selectedIndex < 1) {
		alert("Select a Book");
		document.formname.selBook.focus();
		return;
	}
	if (sFlag === "VouNo") {
		if (document.formname.txtNoFrom.value === "") {
			alert("Enter Voucher No. From ");
			document.formname.txtNoFrom.select();
			return;
		}
		if (document.formname.txtNoTo.value === "") {
			alert("Enter Voucher No. To ");
			document.formname.txtNoTo.select();
			return;
		}
	} else if (sFlag === "VouDate") {
		var sFromDate = controlDate("ctlVouFromDate");
		var sToDate = controlDate("ctlVouToDate");
		var fromDate = parseLegacyDate(sFromDate);
		var toDate = parseLegacyDate(sToDate);
		if (fromDate && toDate && toDate.getTime() < fromDate.getTime()) {
			alert("To Date Should be Greater than From Date");
			return;
		}
	} else if (sFlag === "Amount") {
		if (document.formname.txtGAmount.value === "") {
			alert("Enter From Amount");
			document.formname.txtGAmount.select();
			return;
		}
		if (document.formname.txtLAmount.value === "") {
			alert("Enter To Amount");
			document.formname.txtLAmount.select();
			return;
		}
		if (!isNumeric(document.formname.txtGAmount.value)) {
			alert("Enter Numbers Only");
			document.formname.txtGAmount.select();
			return;
		}
		if (!isNumeric(document.formname.txtLAmount.value)) {
			alert("Enter Numbers Only");
			document.formname.txtLAmount.select();
			return;
		}
		if (numericValue(document.formname.txtGAmount.value) > numericValue(document.formname.txtLAmount.value)) {
			alert("To Amount Should be Greater Than From Amount ");
			document.formname.txtLAmount.value = "";
			document.formname.txtLAmount.select();
			return;
		}
	} else if (sFlag === "AccHead") {
		if (document.formname.SelAccHead.value === "0") {
			alert("Select Account Head");
			document.formname.SelAccHead.focus();
			return;
		}
	}

	document.formname.horgName.value = selectText(document.formname.selUnitId);
	document.formname.hBookName.value = selectText(document.formname.selBook);
	document.formname.hVoucherName.value = selectText(document.formname.selVoucher);
	document.formname.hFromDate.value = controlDate("ctlVouFromDate");
	document.formname.hToDate.value = controlDate("ctlVouToDate");
	document.formname.submit();
}

function clearAccountSelection() {
	document.formname.SelAccHead.selectedIndex = 0;
	document.formname.hAccHead.value = "0";
	setSpan("spAccHead", "");
}

function OptSelection() {
	var criteria = document.formname.optCriteria;
	if (criteria[0].checked) {
		sFlag = criteria[0].value;
		document.formname.txtNoFrom.readOnly = false;
		document.formname.txtNoTo.readOnly = false;
		document.formname.txtGAmount.value = "";
		document.formname.txtLAmount.value = "";
		document.formname.txtGAmount.readOnly = true;
		document.formname.txtLAmount.readOnly = true;
		document.formname.SelAccHead.disabled = true;
		setSpan("spAccHead", "");
	} else if (criteria[1].checked) {
		sFlag = criteria[1].value;
		document.formname.txtNoFrom.value = "";
		document.formname.txtNoTo.value = "";
		document.formname.txtGAmount.value = "";
		document.formname.txtLAmount.value = "";
		document.formname.txtNoFrom.readOnly = true;
		document.formname.txtNoTo.readOnly = true;
		document.formname.txtGAmount.readOnly = true;
		document.formname.txtLAmount.readOnly = true;
		document.formname.SelAccHead.disabled = true;
		setSpan("spAccHead", "");
	} else if (criteria[2].checked) {
		sFlag = criteria[2].value;
		document.formname.txtNoFrom.value = "";
		document.formname.txtNoTo.value = "";
		document.formname.txtNoFrom.readOnly = true;
		document.formname.txtNoTo.readOnly = true;
		document.formname.txtGAmount.readOnly = false;
		document.formname.txtLAmount.readOnly = false;
		document.formname.SelAccHead.disabled = true;
		setSpan("spAccHead", "");
	} else if (criteria[3].checked) {
		sFlag = criteria[3].value;
		document.formname.txtNoFrom.value = "";
		document.formname.txtNoTo.value = "";
		document.formname.txtGAmount.value = "";
		document.formname.txtLAmount.value = "";
		document.formname.txtNoFrom.readOnly = true;
		document.formname.txtNoTo.readOnly = true;
		document.formname.txtGAmount.readOnly = true;
		document.formname.txtLAmount.readOnly = true;
		document.formname.SelAccHead.disabled = false;
	}
}

function SelNew() {
	clearAccountSelection();
	document.formname.txtGAmount.value = "";
	document.formname.txtLAmount.value = "";
	document.formname.txtNoFrom.value = "";
	document.formname.txtNoTo.value = "";
}

function openModernDialog(url, args, features, callback) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	} else {
		window.open(url, "_blank", "height=480,width=420,resizable=no,status=no");
	}
}

function runStringDialog(page, query, done) {
	var features = "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No";
	openModernDialog(page + "?" + query, "", features, function (outValue) {
		var text = trim(outValue);
		var parts;
		if (text === "") {
			return;
		}
		parts = text.split(":");
		if (parts.length === 1) {
			runStringDialog(page, text, done);
			return;
		}
		done(text, parts);
	});
}

function lastAccHeadNode() {
	var nodes = childElements(islandRoot("AccHeadData"));
	return nodes.length ? nodes[nodes.length - 1] : null;
}

function applyAccHeadNode(prefix) {
	var headerNode = lastAccHeadNode();
	if (headerNode) {
		document.formname.hAccHead.value = prefix ? prefix + "?" + attr(headerNode, "No") : attr(headerNode, "No");
		setSpan("spAccHead", attr(headerNode, "Name") + "&nbsp;");
	} else {
		clearAccountSelection();
	}
}

function SelectAccHead() {
	if (document.formname.selUnitId.selectedIndex === 0) {
		alert("Select Organaisation Id");
		clearAccountSelection();
		document.formname.selUnitId.focus();
		return;
	}
	if (document.formname.selVoucher.value === "0") {
		alert("Select Voucher Type");
		document.formname.selVoucher.focus();
		clearAccountSelection();
		return;
	}
	if (document.formname.selBook.value === "S") {
		alert("Select Book");
		document.formname.selBook.focus();
		clearAccountSelection();
		return;
	}

	var sOrgId = document.formname.selUnitId.value;
	var sBookId = document.formname.selVoucher.value;
	var sBookNo = document.formname.selBook.value;

	if (document.formname.SelAccHead.value === "G") {
		runStringDialog("GLHeadSelection.asp", "orgid=" + encodeURIComponent(sOrgId) + "&BookId=" + encodeURIComponent(sBookId) + "&BookNo=" + encodeURIComponent(sBookNo), function (outValue, parts) {
			if (parts.length <= 2 || parts[0] === "-1") {
				return;
			}
			if (typeof window.GetGlHeadXml === "function") {
				window.GetGlHeadXml(outValue);
			}
			applyAccHeadNode("");
		});
	} else {
		var sPartyType = document.formname.SelAccHead.value + "?" + selectText(document.formname.SelAccHead);
		runStringDialog("PartySelection.asp", "orgId=" + encodeURIComponent(sOrgId) + "&Party=" + encodeURIComponent(sPartyType), function (outValue, parts) {
			if (parts.length <= 4 || parts[0] === "-1") {
				return;
			}
			var sPartyName = parts[0];
			var sParCode = parts[1];
			var sParSubType = parts[3];
			var sParTy = parts[4];
			var sRetVal2 = xhrText("XMLGetPayRecCount.asp?orgID=" + encodeURIComponent(sOrgId) + "&ParSubType=" + encodeURIComponent(sParSubType) + "&ParType=" + encodeURIComponent(sParTy) + "&PartyCode=" + encodeURIComponent(sParCode));
			if (trim(sRetVal2) !== "" && typeof window.GetPartyHeadXml === "function") {
				window.GetPartyHeadXml(sParCode, sPartyName, sRetVal2);
			}
			applyAccHeadNode(sPartyType);
		});
	}
}

function SetDate() {
	var sFromYr = "01/04/" + trim(document.formname.hFromYr.value);
	var sToYr = "31/03/" + trim(document.formname.hToYr.value);
	var fromControl = formField("ctlVouFromDate") || document.getElementById("ctlVouFromDate");
	var toControl = formField("ctlVouToDate") || document.getElementById("ctlVouToDate");
	if (fromControl) {
		fromControl.setMinDate ? fromControl.setMinDate(sFromYr) : fromControl.SetMinDate && fromControl.SetMinDate(sFromYr);
		fromControl.setMaxDate ? fromControl.setMaxDate(sToYr) : fromControl.SetMaxDate && fromControl.SetMaxDate(sToYr);
	}
	if (toControl) {
		toControl.setMinDate ? toControl.setMinDate(sFromYr) : toControl.SetMinDate && toControl.SetMinDate(sFromYr);
		toControl.setMaxDate ? toControl.setMaxDate(sToYr) : toControl.SetMaxDate && toControl.SetMaxDate(sToYr);
	}
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetDate()">

<form method="POST" name="formname" action="AppVoucherList.asp">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="hVoucherName" value="">
<input type="hidden" name="horgName" value="">
<input type="hidden" name="hAccHead" value="0">
<input type="hidden" name="hFromDate" value="0">
<input type="hidden" name="hToDate" value="0">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Approve Voucher
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="96">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher List
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Voucher</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" height="2">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="90">Organization</td>
                            <td class="FieldCell">
							<select size="1" name="selUnitId" onchange="popPartType()" class="FormElem" >
							   <OPTION value="0">Select a Unit</option>
							   <%populateOrganizationListDB%>
							</select>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="90">Voucher Type</td>
                            <td class="FieldCell">
							<select size="1" name="selVoucher" class="FormElem" onChange="DisplayBook()">
								<OPTION value="0">Select a Voucher </option>
								<OPTION value="01">Cash Voucher</option>
								<OPTION value="02">Bank Voucher</option>
								<OPTION value="04">Purchase Voucher</option>
								<OPTION value="05">Sales Voucher</option>
								<OPTION value="06">Debit Voucher</option>
								<OPTION value="07">Credit Voucher</option>
								<OPTION value="08">General Voucher</option>
							</select>
                            </td>
                                </tr>
                                <tr>
									<td class="FieldCell" width="90" valign="top">Book</td>
									<td class="FieldCell">
										<select size="1" name="selBook" class="FormElem">
											<option value="S">Select Book</option>
										</select>
									</td>
                                </tr>
                                <tr>
									<td class="FieldCell" width="90" valign="top"></td>
									<td class="FieldCell">
																<table class="ExcelTable" cellSpacing="0" cellPadding="2" border="0">
																<tbody>
																<tr>
																	<td vAlign="center" class="ExcelHeaderCell" align="center">Viewed
                                                                      By</td>
																<td vAlign="center" align="center" class="ExcelHeaderCell">From&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="center" class="ExcelHeaderCell">To&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="center" class="ExcelHeaderCell"></td>
                                                                </tr>
																<tr>
																	<td class="FieldCellSub"><input type="radio" value="VouNo" name="optCriteria" onclick="OptSelection()">
																	Voucher	Number&nbsp;</td>
																<td align="left" class="FieldCellSub"><input class="formelem"  size="11" name="txtNoFrom"></td>
                                                                  <td align="left" class="FieldCellSub"><input class="formelem"  size="11" name="txtNoTo"></td>
                                                                  <td align="left" class="FieldCellSub">&nbsp;</td>
                                                                </tr>
                                                                <tr>
                                                                  <td class="FieldCellSub"><input onclick="OptSelection()" type="radio" value="VouDate" name="optCriteria">
                                                                    Voucher Date</td>
                                                           <td align="left" class="FieldCellSub">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlVouFromDate")
 %>
 </td>
                                                                  <td align="left" class="FieldCellSub">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlVouToDate")
 %>

</td>
                                                                  <td align="left" class="FieldCellSub">

&nbsp;

</td>
                                                                </tr>
                                                                <tr>
                                                                  <td class="FieldCellSub"><input type="radio" onclick="OptSelection()"  value="Amount" name="optCriteria">
                                                                    Amount&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                                                  <td align="left" class="FieldCellSub"><input class="formelem" size="11" Readonly name="txtGAmount"></td>
                                                                  <td align="left" class="FieldCellSub"><input class="formelem" size="11" Readonly name="txtLAmount"></td>
                                                                  <td align="left" class="FieldCellSub"></td>
                                                                </tr>
                                                                <tr>
                                                                  <td class="FieldCellSub"><input type="radio" onclick="OptSelection()" value="AccHead" name="optCriteria">
                                                                    Account Head</td>
                                                                  <td colSpan="3" align="left" class="FieldCellSub">
																	<select class="formelem" disabled OnChange="SelectAccHead()" size="1" name="SelAccHead">
																		  <option value="0">Select Option</option>
																		 <option value="G">General Ledger</option>
																	 </select>
                                                                   </td>
                                                                </tr>
                                                                <tr>
                                                                  <td class="FieldCellSub"></td>
                                                                   <td colSpan="3" class="FieldCellSub"><span id="spAccHead" class="DataOnly"></span>&nbsp;</td>
                                                                </tr>
                                                              </tbody>
                                                            </table>
									</td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="2">
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
															<p align="center">
																<input type="button" value="Next" name="B8" class="ActionButton" onClick="validate()" >
                                                                <input type="reset" value="Reset" name="B9" class="ActionButton" >
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
</BODY>
</HTML>
