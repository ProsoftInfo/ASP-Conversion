<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtReqChequeEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 19, 2003
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
dim sOrgId,sOrgName,sAccCode,sAccName,sReqType,dTransLimit,oDOM
'sOrgId=Request.Form("selUnitId")
'sOrgName=Request.Form("hUnitName")
sOrgId = session("organizationcode")
sOrgName = session("orgshortname")
sReqType=Request.Form("hReqTypeS")


Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

'Response.Write "<p>dTransLimit="&dTransLimit

%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<xml id="PayableData"><Payables ReqType="B"/></xml>
<XML id="PartyData"><Root></Root></XML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT language="javascript">
function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function toNumber(value) {
	var number = Number(String(value == null ? "" : value).replace(/,/g, ""));
	return isFinite(number) ? number : NaN;
}

function formatAmount(value) {
	var number = toNumber(value);
	return isNaN(number) ? "0.00" : number.toFixed(2);
}

function xmlDocument(name) {
	var island = window[name] || document.getElementById(name);
	if (island && island.XMLDocument) {
		return island.XMLDocument;
	}
	if (island && island._doc) {
		return island._doc;
	}
	if (island && island.documentElement) {
		return island;
	}
	return new DOMParser().parseFromString("<Root/>", "text/xml");
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

function loadXmlIsland(name, text) {
	var island = window[name] || document.getElementById(name);
	if (island && typeof island.loadXML === "function") {
		island.loadXML(text);
		return xmlDocument(name);
	}
	return new DOMParser().parseFromString(text || "<Root/>", "text/xml");
}

function childElements(root) {
	var nodes = [];
	for (var i = 0; root && i < root.childNodes.length; i += 1) {
		if (root.childNodes[i].nodeType === 1) {
			nodes.push(root.childNodes[i]);
		}
	}
	return nodes;
}

function attr(node, index) {
	return node && node.attributes && node.attributes[index] ? node.attributes[index].nodeValue : "";
}

function setAttrByIndex(node, index, value) {
	if (node && node.attributes && node.attributes[index]) {
		node.setAttribute(node.attributes[index].name, value == null ? "" : String(value));
	}
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

function clearPayableRows() {
	var table = document.getElementById("tblPayable");
	while (table && table.rows.length > 1) {
		table.deleteRow(1);
	}
}

function resetAccount() {
	document.formname.hAccountCode.value = "";
	document.formname.hAccountName.value = "";
	document.formname.txtPayTo.value = "";
}

function populatePayables(sOrgId, partyCode) {
	var xhr = new XMLHttpRequest();
	var doc;
	var root;
	var nodes;
	var table = document.getElementById("tblPayable");
	xhr.open("GET", "XMLGetPayables.asp?orgId=" + encodeURIComponent(sOrgId) + "&ParCode=" + encodeURIComponent(partyCode), false);
	xhr.send(null);
	if (!xhr.responseText) {
		return;
	}
	doc = loadXmlIsland("PayableData", xhr.responseText);
	root = doc.documentElement;
	nodes = childElements(root);
	for (var i = 0; i < nodes.length; i += 1) {
		var row = table.insertRow(i + 1);
		var docNo = attr(nodes[i], 0);
		InsertCell(row, 1, "", i + 1, "ExcelSerial", "Center", "", 0, 0, 0, 0, "");
		InsertCell(row, 1, "", attr(nodes[i], 1) + "-" + attr(nodes[i], 2), "ExcelDisplayCell", "left", "", 0, 0, 0, 0, "");
		InsertCell(row, 1, "", attr(nodes[i], 3) + "-" + attr(nodes[i], 4), "ExcelDisplayCell", "left", "", 0, 0, 0, 0, "");
		InsertCell(row, 1, "", formatAmount(attr(nodes[i], 5)), "ExcelDisplayCell", "right", "", 0, 0, 0, 0, "");
		InsertCell(row, 1, "", formatAmount(attr(nodes[i], 6)), "ExcelDisplayCell", "right", "", 0, 0, 0, 0, "");
		InsertCell(row, 2, "txtDocAmount" + docNo, "0", "ExcelInputCell", "right", "", 15, 13, 0, 0, "style=\"text-align:right\"");
	}
}

function selAccountHead(objAcc) {
	var sOrgId;
	var sPartyCode;
	var sizeInfo;
	var programName;
	var features;
	var args;
	clearPayableRows();
	if (objAcc.selectedIndex > 0) {
		sPartyCode = objAcc.value + "?" + objAcc.options[objAcc.selectedIndex].text;
		sOrgId = document.formname.hUnitId.value;
		sizeInfo = GetWindowSizeForPopup("2").split(":");
		programName = sizeInfo[0];
		features = "dialogHeight:" + sizeInfo[1] + "px;dialogWidth:" + sizeInfo[2] + "px;Status:No";
		args = window.PartyData || document.getElementById("PartyData");
		runSelectionDialog(programName, "orgid=" + encodeURIComponent(sOrgId) + "&Party=" + encodeURIComponent(sPartyCode), args, features, function (root) {
			var entry = firstEntry(root);
			var partyCode;
			var partyName;
			if (!entry) {
				document.formname.selAcctype.selectedIndex = 0;
				return;
			}
			partyCode = entry.getAttribute("RetField1") || "";
			partyName = entry.getAttribute("RetField0") || "";
			document.formname.hAccountCode.value = sPartyCode + "?" + partyCode;
			document.formname.hAccountName.value = partyName + "&nbsp;";
			document.formname.txtPayTo.value = partyName;
			populatePayables(sOrgId, sPartyCode + "?" + partyCode);
		});
	} else {
		resetAccount();
	}
}

function actionDone() {
	var form = document.formname;
	var doc = xmlDocument("PayableData");
	var root = doc.documentElement;
	var nodes = childElements(root);
	var bFlag = false;
	var dTotal = 0;
	if (form.hAccountCode.value === "") {
		alert("Select Account Head");
		form.selAcctype.focus();
		return;
	}
	for (var i = 0; i < nodes.length; i += 1) {
		var docNo = attr(nodes[i], 0);
		var amountField = form.elements["txtDocAmount" + docNo];
		var amountText = amountField ? amountField.value : "";
		var amount = toNumber(amountText);
		var amtPayable = toNumber(attr(nodes[i], 5));
		var amtPaid = toNumber(attr(nodes[i], 6));
		if (trim(amountText) !== "") {
			if (isNaN(amount)) {
				alert("Enter Numeric Value");
				return;
			}
			if (amount < 0 || amount > 9999999999.99) {
				alert("Amount Should Be > 0 and < 9999999999.99");
				return;
			}
			if (amount > (amtPayable - amtPaid)) {
				alert("Amount is greater than to be paid amount");
				return;
			}
			if (amount > 0) {
				bFlag = true;
				dTotal += amount;
			}
		}
	}
	if (bFlag === false) {
		alert("Request should be created for atleast one Bill ");
		return;
	}
	if (form.hRequestType.value === "C" && dTotal > toNumber(form.hCreditLimit.value)) {
		alert("Cash transcation should not exceed " + form.hCreditLimit.value);
		return;
	}
	if (form.selUserId.selectedIndex === 0) {
		alert("Select Approver ");
		form.selUserId.focus();
		return;
	}
	for (var n = 0; n < nodes.length; n += 1) {
		var nodeDocNo = attr(nodes[n], 0);
		var nodeAmountField = form.elements["txtDocAmount" + nodeDocNo];
		setAttrByIndex(nodes[n], 6, nodeAmountField ? nodeAmountField.value : "0");
	}
	var xhr = new XMLHttpRequest();
	xhr.open("POST", "XMLSave.asp?Mod=CHQ&Name=Payment%20Requestion", false);
	xhr.setRequestHeader("Content-Type", "text/xml");
	xhr.send(new XMLSerializer().serializeToString(doc));
	if (xhr.responseText !== "") {
		alert(xhr.responseText);
	} else {
		form.submit();
	}
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="PmtReqChequeInsert.asp">
<input type="hidden" name="hUnitName" value="">
<input type="hidden" name="hUnitId" value="<%=sOrgId%>">
<input type="hidden" name="hAccountCode" value="">
<input type="hidden" name="hAccountName" value="">
<input type="hidden" name="hCreditLimit" value="<%=dTransLimit%>">


<%if sReqType="A" then %>
<input type="hidden" name="hRequestType" value="B">
<%else%>
<input type="hidden" name="hRequestType" value="C">
<%end if%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="446">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Regular Payment
		<%if sReqType="A" then
			Response.Write "CHEQUE"
		else
			Response.Write "CASH"
		end if
		%>

		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top" height="419">
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
                                                    <table border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                            <td class="FieldCellSub" width="139">Payment Type</td>
                                            <td class="FieldCell"><span class="DataOnly">
                                            <%if sReqType="A" then
												Response.Write "CHEQUE"
											else
												Response.Write "CASH"
											end if
											%>
                                            &nbsp;</span>
                                            </td>
                                                </tr>
                                                <!--<tr>
                                            <td class="FieldCellSub" width="139">Unit</td>
                                            <td class="FieldCell"><span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                                            </td>
                                                </tr>-->
                                                <tr>
                                            <td class="FieldCellSub" width="139">Party Type</td>
                                            <td class="FieldCell">
                                            <select size="1" name="selAcctype" class="FormElem" onChange="selAccountHead(this)">
									   		<option value="S">Select Account Head</option>
									  		 <%populatePartyType(sOrgId)%>
											</select>
                                            </td>
                                                </tr>
                                                <tr>
													<td class="FieldCellSub" width="105">Pay To</td>
													<td class="FieldCell"> <input type="text" name="txtPayTo" size="40" class="FormElem"> </td>
													    </tr>
                                                    </table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top">
									<DIV class=frmBody id=frm1 style="width: 585; height:140;">
                                    <table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center"></td>
                                        <td class="ExcelHeaderCell" align="center">Voucher No- Date</td>
                                        <td class="ExcelHeaderCell" align="center">Bill No - Date</td>
                                        <td class="ExcelHeaderCell" align="center">Bill Amount</td>
                                        <td class="ExcelHeaderCell" align="center">Amount Paid</td>
                                        <td class="ExcelHeaderCell" align="center">Amount To Pay</td>
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
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
                                    <table cellpadding="0" cellspacing="0">
                                <tr>
                            <td class="FieldCell" width="130"> Immediate Approver </td>
                            <td>
									<select size="1" name="selUserId" class="FormElem">
												<option value="0">Immediate Approver</option>
												<%=populateEmployee%>
												    </select>

                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
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
 <input type="button" value="Ok" name="B4" class="ActionButton" onClick="actionDone()">
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
