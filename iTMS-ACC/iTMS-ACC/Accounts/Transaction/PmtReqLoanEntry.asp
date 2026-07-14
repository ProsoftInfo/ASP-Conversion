<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtLoanEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February  1, 2003
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->

<%
dim sOrgId,sOrgName,sAccCode,sAccName,sRequestType

'sOrgId=Request.Form("selUnitId")
'sOrgName=Request.Form("hUnitName")
sOrgId = session("organizationcode")
sOrgName = session("orgshortname")

sAccCode=Request.Form("hAccountCode")
sAccName=Request.Form("hAccountName")
sRequestType=Request.Form("hReqTypeS")
'Response.Write "<p>sRequestType="&sRequestType

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempXMLData"><Root></Root></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<script src="../../scripts/checkdate.js"></script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/VouTransactions.js"></script>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT>
function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function toNumber(value) {
	var number = Number(String(value == null ? "" : value).replace(/,/g, ""));
	return isFinite(number) ? number : NaN;
}

function formField(name) {
	var form = document.formname;
	var lower = String(name).toLowerCase();
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

function islandRoot(name) {
	return xmlRoot(window[name] || document.getElementById(name));
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

function firstEntry(root) {
	var nodes = childElements(root);
	for (var i = 0; i < nodes.length; i += 1) {
		if (String(nodes[i].nodeName).toLowerCase() === "entry") {
			return nodes[i];
		}
	}
	return null;
}

function openModernDialog(url, args, features, callback) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	} else {
		window.open(url, "_blank", "height=500,width=420,resizable=no,status=no");
	}
}

function runSelectionDialog(programName, query, args, features, done) {
	openModernDialog("../../Common/" + programName + "?" + query, args, features, function (outValue) {
		var root = xmlRoot(outValue);
		var action = trim(root && root.getAttribute("Action")).toUpperCase();
		var passQuery = trim(root && root.getAttribute("PassQuery"));
		if (!root || action === "CLOSE") {
			return;
		}
		if (action !== "DONE" && passQuery !== "") {
			runSelectionDialog(programName, passQuery, args, features, done);
			return;
		}
		done(root);
	});
}

function resetAccount() {
	document.formname.hAccountCode.value = "";
	document.formname.hAccountName.value = "";
	document.formname.txtPayTo.value = "";
}

function selAccountHead(objAcc) {
	var selectedText;
	if (objAcc.selectedIndex > 0) {
		selectedText = objAcc.options[objAcc.selectedIndex].text;
		showPartyHead(document.formname.hUnitId.value, objAcc.value + "?" + selectedText);
	} else {
		resetAccount();
	}
}

function showGLHead(sOrgId) {
	var sizeInfo = GetWindowSizeForPopup("5").split(":");
	var programName = sizeInfo[0];
	var features = "dialogHeight:" + sizeInfo[1] + "px;dialogWidth:" + sizeInfo[2] + "px;Status:No";
	var args = window.TempXMLData || document.getElementById("TempXMLData");
	runSelectionDialog(programName, "orgID=" + encodeURIComponent(sOrgId) + "&BookId=00&BookNo=", args, features, function (root) {
		var entry = firstEntry(root);
		var retVal;
		var accRoot;
		var headerNode;
		if (!entry) {
			return;
		}
		retVal = [0, 1, 2, 3, 4, 5, 6].map(function (index) {
			return entry.getAttribute("RetField" + index) || "";
		}).join(":");
		if (typeof window.GetGlHeadXml === "function") {
			window.GetGlHeadXml(retVal);
		}
		accRoot = islandRoot("AccHeadData");
		childElements(accRoot).forEach(function (node) {
			headerNode = node;
		});
		if (headerNode) {
			document.formname.hAccountCode.value = headerNode.getAttribute("No") || "";
			document.formname.hAccountName.value = (headerNode.getAttribute("Name") || "") + "&nbsp;";
			document.formname.txtPayTo.value = headerNode.getAttribute("Name") || "";
		}
	});
}

function showPartyHead(sOrgId, sPartyType) {
	var sizeInfo = GetWindowSizeForPopup("2").split(":");
	var programName = sizeInfo[0];
	var features = "dialogHeight:" + sizeInfo[1] + "px;dialogWidth:" + sizeInfo[2] + "px;Status:No";
	var args = window.PartyData || document.getElementById("PartyData");
	runSelectionDialog(programName, "orgid=" + encodeURIComponent(sOrgId) + "&Party=" + encodeURIComponent(sPartyType), args, features, function (root) {
		var entry = firstEntry(root);
		var partyCode;
		var partyName;
		if (!entry) {
			return;
		}
		partyCode = entry.getAttribute("RetField1") || "";
		partyName = entry.getAttribute("RetField0") || "";
		document.formname.hAccountCode.value = sPartyType + "?" + partyCode;
		document.formname.hAccountName.value = partyName;
		document.formname.txtPayTo.value = partyName;
	});
}

function controlDate(name) {
	var control = formField(name) || document.getElementById(name);
	if (control && typeof control.getDate === "function") {
		return control.getDate();
	}
	if (control && typeof control.GetDate === "function") {
		return control.GetDate();
	}
	return control ? control.value : "";
}

function DisplayTerms(objCounterType) {
	var sDate = controlDate("ctlDate");
	var iInstallNo = document.formname.txtInstallmentNo.value;
	var sType = objCounterType.options[objCounterType.selectedIndex].value;
	document.formname.txtStartDate.value = sDate;
	if (ValidateAmount(document.formname.txtLoanAmount.value) === false) {
		document.formname.txtLoanAmount.select();
		return;
	}
	if (!validateInterest()) {
		objCounterType.selectedIndex = 0;
		return;
	}
	if (trim(iInstallNo) === "") {
		alert("Enter no of Installment");
		document.formname.txtInstallmentNo.focus();
		objCounterType.selectedIndex = 0;
		return;
	}
	if (isNaN(toNumber(iInstallNo))) {
		alert("No of Installment should be number");
		document.formname.txtInstallmentNo.select();
		objCounterType.selectedIndex = 0;
		return;
	}
	ClearTable();
	if (objCounterType.selectedIndex !== 0) {
		for (var j = 1; j <= Number(iInstallNo); j += 1) {
			var row = document.getElementById("tblTerms").insertRow(j);
			InsertCell(row, 1, "", j, "ExcelSerial", "Center", "", 0, 0, 0, 0, "");
			InsertCell(row, 1, "", GetInterval(sDate, sType, j - 1), "ExcelDisplayCell", "left", "", 0, 0, 0, 0, "");
			InsertCell(row, 2, "txtAmount" + j, "", "ExcelInputCell", "", "", 11, 10, 0, 0, "");
			InsertCell(row, 2, "txtPrincipal" + j, "", "ExcelInputCell", "", "", 11, 10, 0, 0, "");
			InsertCell(row, 2, "txtInterst" + j, "", "ExcelInputCell", "", "", 11, 10, 0, 0, "");
		}
	}
}

function ClearTable() {
	var table = document.getElementById("tblTerms");
	while (table && table.rows.length > 1) {
		table.deleteRow(1);
	}
}

function parseLegacyDate(value) {
	var parts = String(value || "").split("/");
	if (parts.length !== 3) {
		return new Date();
	}
	return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
}

function pad2(value) {
	return value < 10 ? "0" + value : String(value);
}

function addMonths(date, months) {
	return new Date(date.getFullYear(), date.getMonth() + months, date.getDate());
}

function GetInterval(sDate, sIntervalType, iInterval) {
	var baseDate = parseLegacyDate(sDate);
	var nextDate = baseDate;
	if (sIntervalType === "M") {
		nextDate = addMonths(baseDate, Number(iInterval));
	} else if (sIntervalType === "Q") {
		nextDate = addMonths(baseDate, Number(iInterval) * 4);
	} else if (sIntervalType === "H") {
		nextDate = addMonths(baseDate, Number(iInterval) * 6);
	} else if (sIntervalType === "Y") {
		nextDate = new Date(baseDate.getFullYear() + Number(iInterval), baseDate.getMonth(), baseDate.getDate());
	}
	return pad2(nextDate.getDate()) + "/" + pad2(nextDate.getMonth() + 1) + "/" + nextDate.getFullYear();
}

function validateInterest() {
	var rateText = document.formname.txtInterstRate.value;
	var rate = toNumber(rateText);
	if (trim(rateText) === "") {
		alert("Enter Interst Rate");
		document.formname.txtInterstRate.select();
		return false;
	}
	if (isNaN(rate)) {
		alert("Interst Rate should be a numeric value");
		document.formname.txtInterstRate.select();
		return false;
	}
	if (rate < 0 || rate > 100) {
		alert("Interst Rate should >0 and <100");
		document.formname.txtInterstRate.select();
		return false;
	}
	return true;
}

function checksubmit() {
	if (document.formname.hAccountCode.value === "") {
		alert("Select Account Head");
		document.formname.selAccType.focus();
		return;
	}
	if (trim(document.formname.txtReason.value) === "") {
		alert("Enter Reason");
		document.formname.txtReason.select();
		return;
	}
	if (ValidateAmount(document.formname.txtLoanAmount.value) === false) {
		document.formname.txtLoanAmount.select();
		return;
	}
	if (!validateInterest()) {
		return;
	}
	if (trim(document.formname.txtInstallmentNo.value) === "") {
		alert("Enter No of Installment");
		document.formname.txtInstallmentNo.select();
		return;
	}
	if (isNaN(toNumber(document.formname.txtInstallmentNo.value))) {
		alert("Installment should be a numeric value");
		document.formname.txtInstallmentNo.select();
		return;
	}
	if (document.formname.selpayTerms.selectedIndex === 0) {
		alert("Select Payment Terms ");
		document.formname.selpayTerms.focus();
		return;
	}
	if (document.formname.selUserId.selectedIndex === 0) {
		alert("Select Approver ");
		document.formname.selUserId.focus();
		return;
	}
	document.formname.B4.disabled = true;
	document.formname.submit();
}

function ValidateAmount(dAmount) {
	var amount = toNumber(dAmount);
	if (trim(dAmount) === "") {
		alert("Amount Cannot be blank");
		return false;
	}
	if (isNaN(amount)) {
		alert("Enter Numeric values for Amount");
		return false;
	}
	if (amount < 1 || amount > 9999999999.99) {
		alert("Amount should be >1 and < 9999999999.99");
		return false;
	}
	return true;
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="PmtReqLoanInsert.asp">
<input type="hidden" name="hFlag" value="<%=sRequestType%>">
<input type="hidden" name="hUnitId" value="<%=sOrgId%>">
<input type="hidden" name="hAccountCode" value="">
<input type="hidden" name="hAccountName" value="">
<input type="hidden" name="txtStartDate" value="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%
			if sRequestType="L" then
				Response.Write "Loan Payment"
			else
				Response.Write "Hire Purchase Payment"
			end if

		%>

		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<!--<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td class="TabCell" valign="bottom" align="center" width="110">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Request Selection</td>
									</tr>
								  </table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="132">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								   <tr>
									  <td align="center">Requisition Details</td>
									</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                &nbsp;
								</td>
                            </tr>
						</table>
					</td>
				</tr>-->
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0">
														<!--<tr>
															<td class=FieldCell width="140"> Unit</td>
															<td width="250">
                                                            <span class="DataOnly"><%=sOrgName%></span>
                                                            </td>
														</tr>-->
														<tr>
															<td class="FieldCell" width="105">Party Type</td>
															<td class="FieldCell">
															    <select size="1" name="selAccType" class="FormElem" onChange="selAccountHead(this)">
									   							 	<option value="A">Select Party Type</option>
									   							 	<!--option value="G">General Ledger</option-->
															   		 <%populatePartyType(sOrgId)%>
															        </select>
															        </td>
															</tr>
															   <tr>
															<td class="FieldCell" width="105">Pay To</td>
															<td class="FieldCell"> <input type="text" name="txtPayTo" size="45" class="FormElem"></span> </td>
															    </tr>
														<tr>
															<td class=FieldCell width="140"> Reason</td>
															<td>
                                                            <input type="text" name="txtReason" size="50" class="FormElem">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140"> Loan Amount</td>
															<td>
                                                            <input type="text" name="txtLoanAmount" size="15" style="text-align:right" class="FormElem">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140"> Interest&nbsp;Rate&nbsp;</td>
															<td>
                                                            <input type="text" name="txtInterstRate" style="text-align:right" size="5" class="FormElem">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140">
                                                            Number of Installments</td>
															<td>
                                                            <input type="text" name="txtInstallmentNo" size="5" style="text-align:right" class="FormElem">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140">
                                                            Starting From</td>
															<td>
															  <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>

                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140"> Repayment
                                                              Terms</td>
															<td>
                                                            <select size="1" name="selpayTerms" class="Formelem" onChange="DisplayTerms(this)">
															 <option value="0">Select Payment Term</option>
                                                              <option value="M">Monthly</option>
                                                              <option value="Q">Quaterly</option>
                                                              <option value="H">Half yearly</option>
                                                              <option value="Y">Yearly</option>
                                                            </select>
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140"> Immediate Approver</td>
															<td class=subCell>
											        <select size="1" name="selUserId" class="FormElem">
											<option value="0">Immediate Approver</option>
											<%=populateEmployee%>
											    </select>
                                                            </td>
														</tr>
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top">
												<DIV class=frmBody id=frm1 style="width: 378; height: 153">
                                                <table border="0" id="tblTerms"cellspacing="1" class="ExcelTable">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="26">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="77">Pay
                                          By</td>
                                        <td class="ExcelHeaderCell" align="center" width="85" valign="top">Amount
                                          to be Paid</td>
                                        <td class="ExcelHeaderCell" align="center" width="85" valign="top">Principal</td>
                                        <td class="ExcelHeaderCell" align="center" width="60">Interest</td>
                                            </tr>

                                                </table>
												</div>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                            <tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
																<input type="button" value="Ok" name="B4" class="ActionButton" onclick="checksubmit()">
																 <input type="reset" value="Reset" name="B1" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
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
