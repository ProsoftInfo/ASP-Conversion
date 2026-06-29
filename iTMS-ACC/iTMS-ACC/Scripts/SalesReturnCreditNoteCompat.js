(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		var lowerName;
		var i;
		if (!frm || !name) {
			return null;
		}
		if (frm.elements[name] || frm[name]) {
			return frm.elements[name] || frm[name];
		}
		lowerName = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === lowerName) {
				return frm.elements[i];
			}
		}
		return null;
	}

	function firstField(item) {
		return item && item.length && !item.tagName ? item[0] : item;
	}

	function fieldValue(name, fallback) {
		var item = field(name);
		var i;
		if (!item) {
			return fallback == null ? "" : fallback;
		}
		if (item.length && !item.tagName) {
			for (i = 0; i < item.length; i += 1) {
				if (item[i].checked) {
					return item[i].value;
				}
			}
			return item[0] ? item[0].value : "";
		}
		return item.value == null ? "" : item.value;
	}

	function setDisabled(name, disabled) {
		var item = firstField(field(name));
		if (item) {
			item.disabled = disabled;
		}
	}

	function focusField(name) {
		var item = firstField(field(name));
		if (item && item.focus) {
			item.focus();
		}
	}

	function upgradeXml() {
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
	}

	function xmlObject(name) {
		var element;
		upgradeXml();
		element = document.getElementById(name);
		return window[name] || document[name] || element && element._itmsXmlIsland || element || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		return object && (object.XMLDocument || object._doc || object) || null;
	}

	function xmlRoot(name) {
		var doc = xmlDocument(name);
		return doc && doc.documentElement || null;
	}

	function clearXmlDocument(doc) {
		while (doc && doc.firstChild) {
			doc.removeChild(doc.firstChild);
		}
	}

	function childElements(node, nodeName) {
		var result = [];
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		var i;
		if (!node) {
			return result;
		}
		for (i = 0; i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function isNode(node, name) {
		return !!node && String(node.nodeName || "").toLowerCase() === String(name).toLowerCase();
	}

	function attr(node, name) {
		var value;
		var lowerName;
		var i;
		if (!node || !node.attributes) {
			return "";
		}
		value = node.getAttribute && node.getAttribute(name);
		if (value != null) {
			return value;
		}
		lowerName = String(name).toLowerCase();
		for (i = 0; i < node.attributes.length; i += 1) {
			if (String(node.attributes[i].name || "").toLowerCase() === lowerName) {
				return node.attributes[i].value || "";
			}
		}
		return "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function nodeText(node) {
		return node ? node.textContent || "" : "";
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var nodes = [];
		var i;
		if (!context) {
			return nodes;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return nodes;
		}
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (i = 0; i < found.snapshotLength; i += 1) {
			nodes.push(found.snapshotItem(i));
		}
		return nodes;
	}

	function requestText(method, url, body) {
		var request = new XMLHttpRequest();
		request.open(method, url, false);
		request.send(body || null);
		return request.responseText || "";
	}

	function getVoucherDate() {
		var control = firstField(field("ctlDate")) || document.getElementById("ctlDate");
		if (control && typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (control && typeof control.getDate === "function") {
			return control.getDate();
		}
		if (control && control.Value != null) {
			return control.Value;
		}
		return control && control.value || "";
	}

	function setDateLimit(control, methodName, value, propertyName) {
		if (!control) {
			return;
		}
		if (typeof control[methodName] === "function") {
			control[methodName](value);
			return;
		}
		if (window.ITMSModernCompat && propertyName) {
			control[propertyName] = window.ITMSModernCompat.toIsoDate(value);
		}
	}

	function setDate() {
		var control = firstField(field("ctlDate")) || document.getElementById("ctlDate");
		var fromDate = "01/04/" + trim(fieldValue("hFromYr"));
		var toDate = "31/03/" + trim(fieldValue("hToYr"));
		setDateLimit(control, "setMinDate", fromDate, "min");
		setDateLimit(control, "SetMinDate", fromDate, "min");
		setDateLimit(control, "setMaxDate", toDate, "max");
		setDateLimit(control, "SetMaxDate", toDate, "max");
	}

	function selectedApprovalValue() {
		var yes = document.querySelector('input[type="radio"][name="optApprove"][value="Y"]');
		return yes && yes.checked ? "Y" : "N";
	}

	function checkFinDate() {
		var finFromText = trim(fieldValue("hFinFrm"));
		var finToText = trim(fieldValue("hFinTo"));
		var finFrom = toNumber(finFromText);
		var finTo = toNumber(finToText);
		var parts = trim(getVoucherDate()).split(/[\/.-]/);
		var period;
		var message;
		if (parts.length < 3) {
			alert("Voucher Date Should Be Between 01/04/" + finFromText.substring(0, 4) + " To 31/03/" + finToText.substring(0, 4));
			return false;
		}
		if (parts[0].length === 4) {
			period = toNumber(parts[0] + (parts[1].length === 1 ? "0" + parts[1] : parts[1]));
		} else {
			period = toNumber(parts[2] + (parts[1].length === 1 ? "0" + parts[1] : parts[1]));
		}
		message = "Voucher Date Should Be Between 01/04/" + finFromText.substring(0, 4) + " To 31/03/" + finToText.substring(0, 4);
		if (period < finFrom || period > finTo) {
			alert(message);
			return false;
		}
		return true;
	}

	function checkNoSer() {
		var passValue;
		var crdr;
		var responseText;
		if (String(fieldValue("hVouCode")) === "04") {
			passValue = fieldValue("selUnitId");
			passValue += ":" + fieldValue("hVouCode");
			passValue += ":" + fieldValue("hCallFrm");
			passValue += ":D";
			passValue += ":" + fieldValue("selBook");
		} else {
			passValue = fieldValue("hOrgid");
			passValue += ":" + fieldValue("hVouCode");
			passValue += ":" + fieldValue("hCallFrm");
			crdr = fieldValue("hVouCRDR");
			passValue += ":" + (trim(crdr) === "" ? "D" : crdr);
			passValue += ":" + fieldValue("hBookCode");
		}
		passValue += ":" + getVoucherDate();
		responseText = trim(requestText("GET", "NoSeriesCheck.asp?sValue=" + passValue));
		if (responseText === "T") {
			return true;
		}
		if (responseText === "F") {
			alert("No Series is Not Defined ");
			return false;
		}
		alert("Error ");
		return false;
	}

	function saveCreditNote(taxRoot) {
		var saleInvoiceNodes;
		var approval = selectedApprovalValue();
		childElements(taxRoot).forEach(function (node) {
			if (isNode(node, "Details")) {
				setAttr(node, "VouDate", getVoucherDate());
			}
		});
		if (approval === "Y" && firstField(field("selUserId")) && firstField(field("selUserId")).selectedIndex === 0) {
			alert("Select Approver ");
			focusField("selUserId");
			return false;
		}
		saleInvoiceNodes = selectNodes(taxRoot, "//SaleInvoice");
		if (saleInvoiceNodes.length !== 0) {
			setAttr(saleInvoiceNodes[0], "Approval", approval);
			setAttr(saleInvoiceNodes[0], "Approver", fieldValue("selUserId"));
			setAttr(saleInvoiceNodes[0], "CrTransNo", fieldValue("hdTransNo"));
		}
		if (!checkFinDate()) {
			return false;
		}
		return finishPost("XMLSave.asp?Mod=CN&Name=Voucher Entry", xmlDocument("TaxData"), "VouCNSalRetAdj.asp");
	}

	function appendNarration(doc, entry) {
		var narration = doc.createElement("Narration");
		narration.textContent = fieldValue("txtNarration");
		entry.appendChild(narration);
	}

	function appendPayRecIfNeeded(doc, entry) {
		var responseText = requestText("GET", "GetGJXML.asp?hTransNo=" + fieldValue("hdTransNo"));
		var values;
		var payRec;
		var docNode;
		if (trim(responseText) === "") {
			return;
		}
		values = responseText.split("#");
		if (toNumber(values[3]) <= toNumber(values[5])) {
			return;
		}
		payRec = doc.createElement("PayRec");
		docNode = doc.createElement("Doc");
		setAttr(docNode, "No", values[0]);
		setAttr(docNode, "InvNo", values[1]);
		setAttr(docNode, "InvDate", values[2]);
		setAttr(docNode, "TransAmount", values[3]);
		setAttr(docNode, "AmtAdjusted", values[4]);
		setAttr(docNode, "AmtToAdjust", fieldValue("txtinvoi", fieldValue("txtInvValue")));
		setAttr(docNode, "DocType", values[6]);
		setAttr(docNode, "AmtToAccount", values[5]);
		setAttr(docNode, "PayableNo", values[8]);
		setAttr(docNode, "AdjType", values[9]);
		payRec.appendChild(docNode);
		entry.appendChild(payRec);
	}

	function appendRecCount(doc, entry) {
		var recCount = doc.createElement("RecCount");
		setAttr(recCount, "Val", "1");
		entry.appendChild(recCount);
	}

	function appendFirstGjEntry(doc, voucherRoot, context, sourceEntry, counter) {
		var entry = doc.createElement("Entry");
		var payRecCount = requestText("GET", "XMLGetPayRecCount.asp?orgID=" + context.partyDetails).split(":");
		setAttr(entry, "No", counter.value);
		setAttr(entry, "CRDR", "C");
		setAttr(entry, "Payto", "0");
		setAttr(entry, "Amount", fieldValue("txtInvValue"));
		setAttr(entry, "AccUnit", context.unitNo);
		setAttr(entry, "AccName", context.unitName);
		setAttr(entry, "TdsAmount", "0.00");
		setAttr(entry, "TDSElgi", "0");
		setAttr(entry, "TdsPercentage", "0");
		setAttr(entry, "PayRecAmount", "0");
		voucherRoot.appendChild(entry);
		childElements(sourceEntry).forEach(function (detail) {
			var accHead;
			if (isNode(detail, "AccHead")) {
				accHead = doc.createElement("AccHead");
				setAttr(accHead, "No", context.partyNo);
				setAttr(accHead, "Pay", payRecCount[0] || "");
				setAttr(accHead, "Rec", payRecCount[1] || "");
				setAttr(accHead, "Name", context.partyName);
				setAttr(accHead, "Type", "P");
				setAttr(accHead, "Adv", payRecCount[2] || "");
				entry.appendChild(accHead);
			}
			appendPayRecIfNeeded(doc, entry);
			appendRecCount(doc, entry);
			appendNarration(doc, entry);
		});
	}

	function appendDebitGjEntry(doc, voucherRoot, context, sourceEntry, counter) {
		var entry = doc.createElement("Entry");
		setAttr(entry, "No", counter.value);
		setAttr(entry, "CRDR", "D");
		setAttr(entry, "Payto", attr(sourceEntry, "PayTo"));
		setAttr(entry, "Amount", fieldValue("txtInvValue"));
		setAttr(entry, "AccUnit", context.unitNo);
		setAttr(entry, "AccName", context.unitName);
		setAttr(entry, "TdsAmount", "0.00");
		setAttr(entry, "TDSElgi", "0");
		setAttr(entry, "TdsPercentage", "0");
		setAttr(entry, "PayRecAmount", "0");
		voucherRoot.appendChild(entry);
		childElements(sourceEntry).forEach(function (detail) {
			var accHead;
			var accountNo;
			if (isNode(detail, "AccHead")) {
				accountNo = attr(detail, "No");
				accHead = doc.createElement("AccHead");
				setAttr(accHead, "No", accountNo);
				setAttr(accHead, "CostCenter", attr(detail, "CostCenter"));
				setAttr(accHead, "Analytical", attr(detail, "Analytical"));
				setAttr(accHead, "Name", attr(detail, "Name"));
				setAttr(accHead, "Type", attr(detail, "Type"));
				setAttr(accHead, "TransFlag", accountNo !== "0" ? "A" : "W");
				entry.appendChild(accHead);
			}
			appendNarration(doc, entry);
		});
	}

	function readHeaderContext(headerNode, voucherRoot) {
		var context = {
			unitNo: "",
			unitName: "",
			bookAccHead: "",
			partyNo: "",
			partyDetails: "",
			partyName: ""
		};
		childElements(headerNode).forEach(function (node) {
			if (isNode(node, "Organization")) {
				context.unitNo = attr(node, "OrgId");
				context.unitName = nodeText(node);
				setAttr(voucherRoot, "UnitNo", context.unitNo);
				setAttr(voucherRoot, "UnitName", context.unitName);
				setAttr(voucherRoot, "BookNo", fieldValue("hBookCode"));
				setAttr(voucherRoot, "BookName", fieldValue("hBookName"));
				setAttr(voucherRoot, "CRDR", "");
			} else if (isNode(node, "Book")) {
				context.bookAccHead = attr(node, "BKAccHead");
			} else if (isNode(node, "Party")) {
				context.partyNo = attr(node, "ParType") + "?" + attr(node, "ParSubType") + "?" + attr(node, "ParType") + "-" + attr(node, "ParSubTypeName") + "?" + attr(node, "ParCode");
				context.partyDetails = context.unitNo + "&ParSubType=" + attr(node, "ParSubType") + "&ParType=" + attr(node, "ParType") + "&PartyCode=" + attr(node, "ParCode");
				context.partyName = nodeText(node);
			}
		});
		return context;
	}

	function buildGjVoucher(taxRoot) {
		var doc = xmlDocument("GJVoucher");
		var voucherRoot;
		var context = {};
		var counter = { value: 0 };
		var firstEntryDone = false;
		if (!doc) {
			return null;
		}
		clearXmlDocument(doc);
		voucherRoot = doc.createElement("voucher");
		doc.appendChild(voucherRoot);
		childElements(taxRoot).forEach(function (section) {
			if (isNode(section, "Header")) {
				context = readHeaderContext(section, voucherRoot);
			} else if (isNode(section, "Details")) {
				setAttr(voucherRoot, "VouDate", getVoucherDate());
				setAttr(voucherRoot, "BookAcchead", context.bookAccHead || "");
				setAttr(voucherRoot, "Approver", fieldValue("selUserId"));
				childElements(section, "Entry").forEach(function (sourceEntry) {
					if (!firstEntryDone) {
						counter.value += 1;
						appendFirstGjEntry(doc, voucherRoot, context, sourceEntry, counter);
						firstEntryDone = true;
					}
					counter.value += 1;
					appendDebitGjEntry(doc, voucherRoot, context, sourceEntry, counter);
				});
			}
		});
		return doc;
	}

	function finishPost(url, doc, action) {
		var responseText = requestText("POST", url, doc);
		var frm;
		if (trim(responseText) !== "") {
			alert(responseText);
			return false;
		}
		setDisabled("B2", true);
		frm = form();
		if (frm) {
			frm.action = action;
			frm.submit();
		}
		return true;
	}

	function saveGeneralJournal(taxRoot) {
		var gjDocument = buildGjVoucher(taxRoot);
		if (!gjDocument || !checkFinDate()) {
			return false;
		}
		return finishPost(
			"XMLSave.asp?Mod=GJ&Name=Voucher Entry",
			gjDocument,
			"VouGenerate.asp?hCallFrom=SR&hReturnQty=" + fieldValue("hTotReturnQty")
		);
	}

	function saveXml() {
		var taxRoot = xmlRoot("TaxData");
		var oldInvoiceValue = toNumber(fieldValue("hTotinvVal"));
		var newInvoiceValue = toNumber(fieldValue("txtInvValue"));
		if (oldInvoiceValue < newInvoiceValue) {
			alert("Returned Invoice Value Should be less than the Invoiced Value ");
			focusField("txtInvValue");
			return false;
		}
		if (!taxRoot) {
			alert("Voucher XML data is not available.");
			return false;
		}
		if (trim(fieldValue("hCallfrom")) === "CR") {
			return saveCreditNote(taxRoot);
		}
		return saveGeneralJournal(taxRoot);
	}

	function enableApproval(source) {
		if (source && source.value === "Y") {
			setDisabled("selUserId", false);
			return;
		}
		if (firstField(field("selUserId"))) {
			firstField(field("selUserId")).selectedIndex = 0;
		}
		setDisabled("selUserId", true);
	}

	function ready(callback) {
		if (document.readyState === "loading") {
			document.addEventListener("DOMContentLoaded", callback);
		} else {
			callback();
		}
	}

	function install() {
		upgradeXml();
		window.SaveXML = saveXml;
		window.EnbApp = enableApproval;
		window.SetDate = setDate;
		window.CheckNoSer = checkNoSer;
		window.CheckFinDate = checkFinDate;
		ready(function () {
			if (window.ITMSModernCompat) {
				window.ITMSModernCompat.init(document);
			}
			setDate();
		});
	}

	window.ITMSSalesReturnCreditNoteCompat = {
		install: install
	};
}(window, document));
