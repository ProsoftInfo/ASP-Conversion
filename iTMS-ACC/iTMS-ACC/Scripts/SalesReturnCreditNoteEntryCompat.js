(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value, decimals) {
		return toNumber(value).toFixed(decimals == null ? 2 : decimals);
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

	function valueOf(name, fallback) {
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

	function setValue(name, value) {
		var item = firstField(field(name));
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function setReadOnly(name, value) {
		var item = firstField(field(name));
		if (item) {
			item.readOnly = !!value;
		}
	}

	function setClassName(name, className) {
		var item = firstField(field(name));
		if (item) {
			item.className = className;
		}
	}

	function byId(id) {
		return document.getElementById(id);
	}

	function setCellClass(id, className) {
		var item = byId(id);
		if (item) {
			item.className = className;
		}
	}

	function setHtml(id, html) {
		var item = byId(id);
		if (item) {
			item.innerHTML = html == null ? "" : String(html);
		}
	}

	function htmlOf(id) {
		var item = byId(id);
		return item ? item.innerHTML : "";
	}

	function textOf(id) {
		var item = byId(id);
		return item ? item.textContent || item.innerText || "" : "";
	}

	function ready(callback) {
		if (document.readyState === "loading") {
			document.addEventListener("DOMContentLoaded", callback);
		} else {
			callback();
		}
	}

	function upgradeModern() {
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		var element;
		upgradeModern();
		element = byId(name);
		return window[name] || document[name] || element && element._itmsXmlIsland || element || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		return object && (object.XMLDocument || object._doc || object) || null;
	}

	function xmlRoot(nameOrObject) {
		var doc;
		if (typeof nameOrObject === "string") {
			doc = xmlDocument(nameOrObject);
			return doc && doc.documentElement || null;
		}
		if (!nameOrObject) {
			return null;
		}
		return nameOrObject.documentElement || nameOrObject._doc && nameOrObject._doc.documentElement || nameOrObject;
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
			if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName || "").toLowerCase() === wanted)) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
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

	function attr(node, nameOrIndex) {
		var lowerName;
		var i;
		var value;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			return node.attributes.item(nameOrIndex) ? node.attributes.item(nameOrIndex).value || "" : "";
		}
		value = node.getAttribute && node.getAttribute(nameOrIndex);
		if (value != null) {
			return value;
		}
		lowerName = String(nameOrIndex).toLowerCase();
		for (i = 0; i < node.attributes.length; i += 1) {
			if (String(node.attributes[i].name || "").toLowerCase() === lowerName) {
				return node.attributes[i].value || "";
			}
		}
		return "";
	}

	function setAttr(node, nameOrIndex, value) {
		var attribute;
		if (!node || !node.attributes) {
			return;
		}
		if (typeof nameOrIndex === "number") {
			attribute = node.attributes.item(nameOrIndex);
			if (attribute) {
				attribute.value = value == null ? "" : String(value);
			}
			return;
		}
		if (node.setAttribute) {
			node.setAttribute(nameOrIndex, value == null ? "" : String(value));
		}
	}

	function appendTextElement(doc, parent, name, value) {
		var node = doc.createElement(name);
		node.textContent = value == null ? "" : String(value);
		parent.appendChild(node);
		return node;
	}

	function requestText(method, url, body) {
		var request = new XMLHttpRequest();
		request.open(method, url, false);
		request.send(body || null);
		return request.responseText || "";
	}

	function getVoucherDate() {
		var control = firstField(field("ctlDate")) || byId("ctlDate");
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

	function setVoucherDate(value) {
		var control = firstField(field("ctlDate")) || byId("ctlDate");
		if (!control || trim(value) === "") {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (window.ITMSModernCompat) {
			control.value = window.ITMSModernCompat.toIsoDate(value);
		} else {
			control.value = value;
		}
	}

	function applyDateLimits() {
		var control = firstField(field("ctlDate")) || byId("ctlDate");
		var fromDate = valueOf("hFromDate");
		var toDate = valueOf("hToDate");
		var today = new Date();
		var todayIso;
		var toIso;
		if (!control || !fromDate || !toDate || !window.ITMSModernCompat) {
			return;
		}
		if (typeof control.SetMinDate === "function") {
			control.SetMinDate(fromDate);
		}
		todayIso = window.ITMSModernCompat.toIsoDate(today);
		toIso = window.ITMSModernCompat.toIsoDate(toDate);
		if (toIso && toIso < todayIso) {
			if (typeof control.SetMaxDate === "function") {
				control.SetMaxDate(toDate);
			}
			setVoucherDate(toDate);
		} else {
			if (typeof control.SetMaxDate === "function") {
				control.SetMaxDate(today);
			}
			setVoucherDate(today);
		}
	}

	function isAmendmentPage() {
		var frm = form();
		return !!xmlObject("InvData") || !!field("hCrTransNo") || !!(frm && /Amd/i.test(frm.action || ""));
	}

	function pageName() {
		return String(window.location.pathname || "").split(/[\/\\]/).pop().toLowerCase();
	}

	function isInvoiceAmendmentPage() {
		return /^vou(?:cn|dn)(?:pur|sal)invamd\.asp$/i.test(pageName());
	}

	function shouldResetFormOnModeApply() {
		return !/^voudnpurinvamd\.asp$/i.test(pageName());
	}

	function invoiceNode(root) {
		return selectNodes(root, "//SaleInvoice")[0] || selectNodes(root, "//PurInvoice")[0] || null;
	}

	function invoiceTransactionAttribute(node) {
		var nodeName = node && String(node.nodeName || "").toLowerCase();
		if (/^voundotherinventry\.asp$/i.test(pageName())) {
			return "CrTransNo";
		}
		if (/^voudnsalinventry\.asp$/i.test(pageName())) {
			return "PurTransNo";
		}
		if (nodeName === "purinvoice" && !/^voudnpurinvamd\.asp$/i.test(pageName())) {
			return "PurTransNo";
		}
		return "SalTrNo";
	}

	function invoiceTotalValue(taxDetails) {
		return attr(taxDetails, "InvoiceVlaue") || attr(taxDetails, "InvoiceValue") || attr(taxDetails, "NettValue") || attr(taxDetails, 0);
	}

	function isDirectNoteEntryPage() {
		var name = pageName();
		if (/^vou(?:dnother|dnsal)inventry\.asp$/i.test(name)) {
			return trim(valueOf("hCallFromDebit")).toUpperCase() === "DR";
		}
		if (/^voucnpurinventry\.asp$/i.test(name)) {
			return trim(valueOf("hCallFrom")).toUpperCase() === "CR";
		}
		return trim(valueOf("hCallFromVoucher")).toUpperCase() === "CR";
	}

	function isDebitGeneralJournalPage() {
		return /^vou(?:dnother|dnsal)inventry\.asp$/i.test(pageName());
	}

	function crValueCellId() {
		return byId("tDrVal") ? "tDrVal" : "tCrValue";
	}

	function setCrNoteEditable(editable) {
		setReadOnly("txtCrNoteValue", !editable);
		setClassName("txtCrNoteValue", editable ? "FormElem" : "FormElemRead");
		setCellClass(crValueCellId(), editable ? "ExcelInputCell" : "ExcelDisplayCell");
	}

	function setRowMode(index, mode) {
		var editable = {
			qty: mode === "Q",
			rate: mode === "R",
			dis: mode === "D"
		};
		setClassName("txtQty" + index, editable.qty ? "FormElem" : "FormElemRead");
		setCellClass("tQty" + index, editable.qty ? "ExcelInputCell" : "ExcelDisplayCell");
		setClassName("txtRate" + index, editable.rate ? "FormElem" : "FormElemRead");
		setCellClass("tRate" + index, editable.rate ? "ExcelInputCell" : "ExcelDisplayCell");
		setClassName("txtDis" + index, editable.dis ? "FormElem" : "FormElemRead");
		setCellClass("tDis" + index, editable.dis ? "ExcelInputCell" : "ExcelDisplayCell");
		setReadOnly("txtDis" + index, !editable.dis);
		setReadOnly("txtQty" + index, !editable.qty);
		setReadOnly("txtRate" + index, !editable.rate);
	}

	function invoiceParts() {
		var parts = trim(textOf("InoviceNo")).split("-");
		return {
			no: trim(parts[0] || ""),
			date: trim(parts[1] || "")
		};
	}

	function updateNarration(mode, totalQty) {
		var invoice = invoiceParts();
		var text = "CR Note for " + invoice.no + " " + invoice.date;
		if (trim(mode) === "Q") {
			text += " Against Quantity: " + totalQty;
		}
		if (trim(valueOf("hInvCallFrom")) === "SR") {
			text = "CR Note for " + invoice.no + " " + invoice.date + " Sales Return Qty: " + totalQty;
		}
		if (field("txtNarration")) {
			setValue("txtNarration", text);
		}
	}

	function setRetVal(source, callType) {
		var select = source || field("SelCrAgain");
		var mode = select ? select.value : "";
		var rowCount = Math.max(0, Math.floor(toNumber(valueOf("hRowVal", 0))));
		var totalQty = "";
		var frm = form();
		var i;
		if (String(callType) === "1") {
			if (frm && shouldResetFormOnModeApply()) {
				frm.reset();
				select = field("SelCrAgain");
			}
			if (select) {
				for (i = 0; i < select.options.length; i += 1) {
					if (select.options[i].value === mode) {
						select.selectedIndex = i;
						break;
					}
				}
			}
		} else if (String(callType) === "2" && isAmendmentPage()) {
			resetCrVal();
		}
		for (i = 1; i <= rowCount; i += 1) {
			if (mode === "Q" || mode === "R" || mode === "D" || mode === "0") {
				setRowMode(i, mode);
				setCrNoteEditable(false);
				if (mode === "Q" || mode === "R" || mode === "D") {
					setValue("txtCrNoteValue", formatNumber(valueOf("txtInvValue", 0), 2));
				}
			} else if (mode === "A") {
				setRowMode(i, "0");
				setCrNoteEditable(true);
				resetTax();
			}
			if (mode === "Q") {
				totalQty += valueOf("txtQty" + i, "");
			}
		}
		if (!isAmendmentPage()) {
			updateNarration(mode, totalQty);
		}
	}

	function resetTax() {
		var root = xmlRoot("TaxData");
		var taxRoot = selectNodes(root, "//TaxDetails")[0];
		if (!taxRoot) {
			return;
		}
		childElements(taxRoot).forEach(function (node) {
			var catCode = attr(node, 0);
			var taxCode = attr(node, 1);
			setAttr(node, 5, "0.00");
			setValue("txtTaxValue" + catCode + taxCode, "0.00");
		});
		setValue("txtInvValue", "0.00");
		setAttr(taxRoot, 0, "0.00");
		setAttr(taxRoot, 1, "0.00");
		setAttr(taxRoot, 2, "0.00");
	}

	function nodeByNo(nodes, number) {
		var i;
		for (i = 0; i < nodes.length; i += 1) {
			if (String(attr(nodes[i], "No")) === String(number)) {
				return nodes[i];
			}
		}
		return null;
	}

	function resetCrVal() {
		var invRoot = xmlRoot("InvData");
		var root = xmlRoot("TaxData");
		var invEntries = selectNodes(invRoot, "//Entry");
		var taxEntries = selectNodes(root, "//Entry");
		var i;
		var rowNo;
		var invTaxNodes;
		setRetVal(field("SelCrAgain"), "1");
		for (i = 0; i < invEntries.length; i += 1) {
			rowNo = i + 1;
			setValue("txtqty" + rowNo, formatNumber(attr(invEntries[i], "Qty"), 2));
			setValue("txtRate" + rowNo, formatNumber(attr(invEntries[i], "Rate"), 2));
			setValue("txtDis" + rowNo, formatNumber(attr(invEntries[i], "DisAmount"), 2));
			setValue("txtAmount" + rowNo, formatNumber(attr(invEntries[i], "Amount"), 2));
			(function (target) {
				if (target) {
					setAttr(target, "Qty", formatNumber(attr(invEntries[i], "Qty"), 2));
					setAttr(target, "Rate", formatNumber(attr(invEntries[i], "Rate"), 2));
					setAttr(target, "DisPer", formatNumber(attr(invEntries[i], "DisPer"), 2));
					setAttr(target, "Amount", formatNumber(attr(invEntries[i], "Amount"), 2));
				}
			}(nodeByNo(taxEntries, rowNo)));
		}
		(function (details) {
			if (details) {
				setValue("txtTotal", formatNumber(attr(details, "BasicValue"), 2));
			}
		}(selectNodes(invRoot, "//Details")[0]));
		(function (taxDetails) {
			if (taxDetails) {
				setValue("txtInvValue", formatNumber(invoiceTotalValue(taxDetails), 2));
			}
		}(selectNodes(invRoot, "//TaxDetails")[0]));
		invTaxNodes = selectNodes(invRoot, "//Tax");
		for (i = 0; i < invTaxNodes.length; i += 1) {
			setValue(
				"txtTaxValue" + attr(invTaxNodes[i], "CatCode") + attr(invTaxNodes[i], "TaxCode"),
				formatNumber(attr(invTaxNodes[i], "TaxAmount"), 2)
			);
		}
	}

	function selectedApprovalValue() {
		var yes = document.querySelector('input[type="radio"][name="optApprove"][value="Y"]');
		return yes && yes.checked ? "Y" : "N";
	}

	function enableApproval(source) {
		if (source && source.value === "Y") {
			if (firstField(field("selUserId"))) {
				firstField(field("selUserId")).disabled = false;
			}
			return;
		}
		if (firstField(field("selUserId"))) {
			firstField(field("selUserId")).selectedIndex = 0;
			firstField(field("selUserId")).disabled = true;
		}
	}

	function checkFinDate() {
		var finFromText = trim(valueOf("hFinFrm"));
		var finToText = trim(valueOf("hFinTo"));
		var parts;
		var period;
		var message;
		if (!finFromText || !finToText) {
			return true;
		}
		parts = trim(getVoucherDate()).split(/[\/.-]/);
		message = "Voucher Date Should Be Between 01/04/" + finFromText.substring(0, 4) + " To 31/03/" + finToText.substring(0, 4);
		if (parts.length < 3) {
			alert(message);
			return false;
		}
		if (parts[0].length === 4) {
			period = toNumber(parts[0] + (parts[1].length === 1 ? "0" + parts[1] : parts[1]));
		} else {
			period = toNumber(parts[2] + (parts[1].length === 1 ? "0" + parts[1] : parts[1]));
		}
		if (period < toNumber(finFromText) || period > toNumber(finToText)) {
			alert(message);
			return false;
		}
		return true;
	}

	function checkNoSer() {
		var passValue;
		var responseText;
		if (String(valueOf("hVouCode")) === "04") {
			passValue = valueOf("selUnitId") + ":" + valueOf("hVouCode") + ":" + valueOf("hCallFrm") + ":D:" + valueOf("selBook");
		} else {
			passValue = valueOf("hOrgid", valueOf("hOrgId")) + ":" + valueOf("hVouCode") + ":" + valueOf("hCallFrm") + ":";
			passValue += trim(valueOf("hVouCRDR")) === "" ? "D" : valueOf("hVouCRDR");
			passValue += ":" + valueOf("hBookCode");
		}
		passValue += ":" + getVoucherDate();
		responseText = trim(requestText("GET", "NoSeriesCheck.asp?sValue=" + passValue));
		if (responseText === "T") {
			return true;
		}
		alert(responseText === "F" ? "No Series is Not Defined " : "Error ");
		return false;
	}

	function updateVoucherDate(root) {
		childElements(root).forEach(function (node) {
			if (String(node.nodeName).toLowerCase() === "details") {
				setAttr(node, "VouDate", getVoucherDate());
			}
		});
	}

	function applyApproval(root) {
		var invoice = invoiceNode(root);
		var approval = selectedApprovalValue();
		if (approval === "Y" && firstField(field("selUserId")) && firstField(field("selUserId")).selectedIndex === 0) {
			alert("Select Approver ");
			firstField(field("selUserId")).focus();
			return false;
		}
		if (invoice) {
			setAttr(invoice, "Approval", approval);
			setAttr(invoice, "Approver", valueOf("selUserId"));
			setAttr(invoice, invoiceTransactionAttribute(invoice), valueOf("hdTransNo"));
		}
		return true;
	}

	function shouldUpdateEntryAmounts() {
		if (valueOf("SelCrAgain") !== "A") {
			return false;
		}
		return !/^vou(?:cnsal|dnpur)invamd\.asp$/i.test(pageName());
	}

	function applyQualityAmounts(root) {
		var entries;
		var i;
		if (!shouldUpdateEntryAmounts()) {
			return;
		}
		entries = selectNodes(root, "//Entry");
		for (i = 0; i < entries.length; i += 1) {
			setAttr(entries[i], "Amount", valueOf("txtAmount" + (i + 1)));
		}
	}

	function invoiceEntryElementName(root) {
		return selectNodes(root, "//PurInvoice")[0] ? "PurchaseInvoiceEntry" : "SalesInvoiceEntry";
	}

	function applyDirectNoteAttributes(root) {
		var header;
		var invoice;
		if (/^voundotherinventry\.asp$/i.test(pageName())) {
			header = selectNodes(root, "//Header")[0];
			if (header) {
				setAttr(header, "PayableAt", valueOf("SelCrAgain"));
			}
		}
		if (/^voucnpurinventry\.asp$/i.test(pageName())) {
			invoice = invoiceNode(root);
			if (invoice) {
				setAttr(invoice, "CRNoteType", valueOf("SelCrAgain"));
			}
		}
	}

	function shouldSaveTaxDetailsTotal() {
		if (/^voucnpurinventry\.asp$/i.test(pageName())) {
			return valueOf("SelCrAgain") === "A";
		}
		return !isAmendmentPage() || valueOf("SelCrAgain") === "A";
	}

	function shouldZeroTaxForAmountMode() {
		if (isAmendmentPage() || /^vou(?:dnother|dnsal)inventry\.asp$/i.test(pageName())) {
			return false;
		}
		return valueOf("SelCrAgain") === "A";
	}

	function directNoteSaveUrl() {
		if (/^vou(?:dnother|cnpur)inventry\.asp$/i.test(pageName())) {
			return "XMLSave.asp?Mod=DN&Name=Voucher Entry";
		}
		return "XMLSave.asp?Mod=CN&Name=Voucher Entry";
	}

	function directNoteSubmitAction() {
		var name = pageName();
		if (/^voundotherinventry\.asp$/i.test(name)) {
			return "VouDNPurInvGenerate.asp";
		}
		if (/^voudnsalinventry\.asp$/i.test(name)) {
			return "VouDNSalInvUpdate.asp";
		}
		if (/^voucnpurinventry\.asp$/i.test(name)) {
			return "VouCNPurInvUpdate.asp";
		}
		return trim(valueOf("hInvCallFrom")) === "SR" ? "VouCNSalRetAdj.asp" : "VouCNOthInvGenerate.asp";
	}

	function disableNextButton() {
		var button = firstField(field("B2"));
		if (button) {
			button.disabled = true;
		}
	}

	function applyDiscountAmounts(root) {
		var entries;
		var i;
		var discount;
		if (!/^voudnpurinvamd\.asp$/i.test(pageName()) || valueOf("SelCrAgain") !== "D") {
			return true;
		}
		entries = selectNodes(root, "//Entry");
		for (i = 0; i < entries.length; i += 1) {
			discount = valueOf("txtDis" + (i + 1));
			if (trim(discount) === "" || isNaN(Number(String(discount).replace(/,/g, "")))) {
				alert("Enter Only Numeric Values");
				if (field("txtDis" + (i + 1)) && field("txtDis" + (i + 1)).focus) {
					field("txtDis" + (i + 1)).focus();
				}
				return false;
			}
			setAttr(entries[i], "Amount", valueOf("txtAmount" + (i + 1)));
			setAttr(entries[i], "DisPer", discount);
			setAttr(entries[i], "DisAmount", discount);
		}
		return true;
	}

	function applyBookAmendment(root) {
		var book;
		var select;
		var selectedOption;
		var details;
		if (!/^voudnpurinvamd\.asp$/i.test(pageName()) || String(valueOf("hAmdTy")) !== "A") {
			return;
		}
		book = selectNodes(root, "//Book")[0];
		select = firstField(field("selBook"));
		if (book && select) {
			selectedOption = select.options[select.selectedIndex] || null;
			setAttr(book, "BookId", select.value);
			book.textContent = selectedOption ? selectedOption.text : select.value;
		}
		details = selectNodes(root, "//Details")[0];
		if (details) {
			setAttr(details, "VouDate", getVoucherDate());
		}
	}

	function amendmentSaveUrl() {
		return isInvoiceAmendmentPage() ? "XMLSave.asp?Mod=DN&Name=Voucher AMD" : "XMLSave.asp?Mod=CNAmd&Name=Voucher Entry";
	}

	function amendmentSubmitAction() {
		if (/^voudnpurinvamd\.asp$/i.test(pageName()) && String(valueOf("hAmdTy")) === "A") {
			return "AmdAccDbNtGenerate.asp";
		}
		return "";
	}

	function applyAccountHead(root) {
		var accHead;
		var description;
		if (valueOf("SelAccountHd") !== "G") {
			return;
		}
		accHead = valueOf("hCrAccHead");
		description = htmlOf("spAccHead");
		selectNodes(root, "//AccHead").forEach(function (node) {
			setAttr(node, "No", accHead);
			setAttr(node, "Name", description);
		});
		selectNodes(root, "//Tax").forEach(function (node) {
			setAttr(node, "AccHead", accHead);
		});
	}

	function finishPost(url, doc, action, beforeSubmit) {
		var responseText = requestText("POST", url, doc);
		var frm;
		if (trim(responseText) !== "") {
			alert(responseText);
			return false;
		}
		frm = form();
		if (frm) {
			if (action) {
				frm.action = action;
			}
			if (beforeSubmit) {
				beforeSubmit(frm);
			}
			frm.submit();
		}
		return true;
	}

	function saveCreditNote(root) {
		var doc = xmlDocument("TaxData");
		var taxDetails;
		updateVoucherDate(root);
		applyBookAmendment(root);
		applyDirectNoteAttributes(root);
		if (isAmendmentPage()) {
			(function (narration) {
				if (narration) {
					narration.textContent = valueOf("txtNarration");
				} else {
					appendTextElement(doc, root, "Narration", valueOf("txtNarration"));
				}
			}(selectNodes(root, "//Voucher/Narration")[0]));
		} else {
			appendTextElement(doc, root, "Narration", valueOf("txtNarration"));
			if (field("hInvoiceNo")) {
				(function (entry) {
					setAttr(entry, "InvoiceNo", valueOf("hInvoiceNo"));
				}(appendTextElement(doc, root, invoiceEntryElementName(root), "")));
			}
		}
		if (!applyApproval(root)) {
			return false;
		}
		taxDetails = selectNodes(root, "//TaxDetails")[0];
		if (taxDetails && shouldSaveTaxDetailsTotal()) {
			setAttr(taxDetails, 0, valueOf("txtCrNoteValue"));
		}
		applyQualityAmounts(root);
		if (!applyDiscountAmounts(root)) {
			return false;
		}
		applyAccountHead(root);
		if (shouldZeroTaxForAmountMode()) {
			selectNodes(root, "//Tax").forEach(function (node) {
				setAttr(node, "TaxAmount", "0.00");
			});
		}
		if (!isAmendmentPage() && !checkFinDate()) {
			return false;
		}
		if (isAmendmentPage()) {
			if (isInvoiceAmendmentPage() && !checkFinDate()) {
				return false;
			}
			return finishPost(amendmentSaveUrl(), doc, amendmentSubmitAction());
		}
		return finishPost(
			directNoteSaveUrl(),
			doc,
			directNoteSubmitAction(),
			/^voucnpurinventry\.asp$/i.test(pageName()) ? disableNextButton : null
		);
	}

	function appendNarration(doc, entry, text) {
		appendTextElement(doc, entry, "Narration", text);
	}

	function appendPayRecIfNeeded(doc, entry) {
		var responseText = requestText("GET", "GetGJXML.asp?hTransNo=" + valueOf("hdTransNo"));
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
		setAttr(docNode, "AmtToAdjust", valueOf("txtCrNoteValue"));
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

	function firstGjCrDr() {
		return isDebitGeneralJournalPage() ? "D" : "C";
	}

	function secondGjCrDr() {
		return isDebitGeneralJournalPage() ? "C" : "D";
	}

	function taxGjCrDr(amount) {
		if (isDebitGeneralJournalPage()) {
			return toNumber(amount) < 0 ? "D" : "C";
		}
		return toNumber(amount) < 0 ? "C" : "D";
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
			if (String(node.nodeName).toLowerCase() === "organization") {
				context.unitNo = attr(node, "OrgId");
				context.unitName = node.textContent || "";
				setAttr(voucherRoot, "UnitNo", context.unitNo);
				setAttr(voucherRoot, "UnitName", context.unitName);
				setAttr(voucherRoot, "BookNo", valueOf("hBookCode"));
				setAttr(voucherRoot, "BookName", valueOf("hBookName"));
				setAttr(voucherRoot, "CRDR", "");
			} else if (String(node.nodeName).toLowerCase() === "book") {
				context.bookAccHead = attr(node, "BKAccHead");
			} else if (String(node.nodeName).toLowerCase() === "party") {
				context.partyNo = attr(node, "ParType") + "?" + attr(node, "ParSubType") + "?" + attr(node, "ParType") + "-" + attr(node, "ParSubTypeName") + "?" + attr(node, "ParCode");
				context.partyDetails = context.unitNo + "&ParSubType=" + attr(node, "ParSubType") + "&ParType=" + attr(node, "ParType") + "&PartyCode=" + attr(node, "ParCode");
				context.partyName = node.textContent || "";
			}
		});
		return context;
	}

	function appendFirstGjEntry(doc, voucherRoot, context, sourceEntry, counter) {
		var entry = doc.createElement("Entry");
		var payRecCount = requestText("GET", "XMLGetPayRecCount.asp?orgID=" + context.partyDetails).split(":");
		setAttr(entry, "No", counter.value);
		setAttr(entry, "CRDR", firstGjCrDr());
		setAttr(entry, "Payto", "0");
		setAttr(entry, "Amount", valueOf("txtCrNoteValue"));
		setAttr(entry, "AccUnit", context.unitNo);
		setAttr(entry, "AccName", context.unitName);
		setAttr(entry, "TdsAmount", "0.00");
		setAttr(entry, "TDSElgi", "0");
		setAttr(entry, "TdsPercentage", "0");
		setAttr(entry, "PayRecAmount", "0");
		voucherRoot.appendChild(entry);
		childElements(sourceEntry).forEach(function (detail) {
			var accHead;
			if (String(detail.nodeName).toLowerCase() === "acchead") {
				accHead = doc.createElement("AccHead");
				setAttr(accHead, "No", context.partyNo);
				setAttr(accHead, "Pay", payRecCount[0] || "");
				setAttr(accHead, "Rec", payRecCount[1] || "");
				setAttr(accHead, "Name", context.partyName);
				setAttr(accHead, "Type", "P");
				setAttr(accHead, "Adv", payRecCount[2] || "");
				entry.appendChild(accHead);
			}
		});
		appendPayRecIfNeeded(doc, entry);
		appendRecCount(doc, entry);
		appendNarration(doc, entry, valueOf("txtNarration"));
	}

	function appendSecondGjEntry(doc, voucherRoot, context, sourceEntry, counter, rowIndex) {
		var entry = doc.createElement("Entry");
		var ratePer = toNumber(attr(sourceEntry, "RatePer")) || 1;
		var narration = attr(sourceEntry, "PayTo") + "&" + valueOf("txtQty" + rowIndex, valueOf("txtqty" + rowIndex)) + "&" + attr(sourceEntry, "UOMValue") + "&" + (toNumber(valueOf("txtRate" + rowIndex)) / ratePer);
		setAttr(entry, "No", counter.value);
		setAttr(entry, "CRDR", secondGjCrDr());
		setAttr(entry, "Payto", "");
		setAttr(entry, "Amount", valueOf("txtTotal"));
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
			if (String(detail.nodeName).toLowerCase() === "acchead") {
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
		});
		appendNarration(doc, entry, narration);
	}

	function appendTaxGjEntries(doc, voucherRoot, taxRoot, context, counter) {
		childElements(taxRoot).forEach(function (taxNode) {
			var catCode = attr(taxNode, 0);
			var taxCode = attr(taxNode, 1);
			var taxAmount = valueOf("txtTaxValue" + catCode + taxCode, "");
			var accHead = attr(taxNode, 6);
			var entry;
			var accNode;
			var secondEntry;
			if (accHead !== "0") {
				if (trim(taxAmount) === "" || toNumber(taxAmount) === 0) {
					return;
				}
				entry = doc.createElement("Entry");
				counter.value += 1;
				setAttr(entry, "No", counter.value);
				setAttr(entry, "CRDR", taxGjCrDr(taxAmount));
				setAttr(entry, "Payto", "");
				setAttr(entry, "Amount", Math.abs(Math.round(toNumber(taxAmount) * 100) / 100));
				setAttr(entry, "AccUnit", context.unitNo);
				setAttr(entry, "AccName", context.unitName);
				setAttr(entry, "TdsAmount", "0.00");
				setAttr(entry, "TDSElgi", "0");
				setAttr(entry, "TdsPercentage", "0");
				setAttr(entry, "PayRecAmount", "0");
				voucherRoot.appendChild(entry);
				accNode = doc.createElement("AccHead");
				setAttr(accNode, "No", accHead);
				setAttr(accNode, "CostCenter", "0");
				setAttr(accNode, "Analytical", "0");
				setAttr(accNode, "Name", "");
				setAttr(accNode, "Type", "G");
				setAttr(accNode, "TransFlag", "A");
				entry.appendChild(accNode);
				appendNarration(doc, entry, taxNode.textContent || "");
			} else {
				secondEntry = childElements(voucherRoot, "Entry").filter(function (node) {
					return attr(node, "No") === "2";
				})[0];
				if (secondEntry) {
					setAttr(secondEntry, "Amount", toNumber(attr(secondEntry, "Amount")) + toNumber(attr(taxNode, 5)));
				}
			}
		});
	}

	function savePurchaseCreditGeneralJournal(root) {
		var doc = xmlDocument("GJVoucher");
		var voucherRoot;
		var handledEntry = false;
		if (!doc) {
			alert("Voucher XML data is not available.");
			return false;
		}
		if (isDebitGeneralJournalPage()) {
			setValue("hVouCode", "08");
		}
		clearXmlDocument(doc);
		voucherRoot = doc.createElement("voucher");
		doc.appendChild(voucherRoot);
		childElements(root).forEach(function (section) {
			if (String(section.nodeName).toLowerCase() === "header") {
				childElements(section).forEach(function (node) {
					if (String(node.nodeName).toLowerCase() === "organization") {
						setAttr(voucherRoot, "UnitNo", attr(node, "OrgId"));
						setAttr(voucherRoot, "UnitName", node.textContent || "");
						setAttr(voucherRoot, "BookNo", valueOf("hBookCode"));
						setAttr(voucherRoot, "BookName", valueOf("hBookName"));
						setAttr(voucherRoot, "CRDR", "D");
					} else if (String(node.nodeName).toLowerCase() === "party") {
						setAttr(voucherRoot, "PartyCode", attr(node, "ParType") + "?" + attr(node, "ParSubType") + "?" + attr(node, "ParSubTypeName") + "?" + attr(node, "ParCode"));
						setAttr(voucherRoot, "Approver", valueOf("selUserId"));
						setAttr(voucherRoot, "PartyName", node.textContent || "");
					} else if (String(node.nodeName).toLowerCase() === "purinvoice") {
						setAttr(voucherRoot, "InvNo", attr(node, "PurInvNo"));
						setAttr(voucherRoot, "InvDate", attr(node, "PurInvDate"));
					}
				});
			} else if (String(section.nodeName).toLowerCase() === "details") {
				setAttr(voucherRoot, "VouDate", getVoucherDate());
				childElements(section, "Entry").forEach(function (sourceEntry) {
					var entry;
					if (handledEntry) {
						return;
					}
					entry = doc.createElement("Entry");
					setAttr(entry, "No", "1");
					setAttr(entry, "CRDR", "C");
					setAttr(entry, "Payto", attr(sourceEntry, "PayTo"));
					setAttr(entry, "Amount", valueOf("txtCrNoteValue"));
					setAttr(entry, "AccUnit", "");
					setAttr(entry, "AccName", "");
					setAttr(entry, "TDSAmount", "");
					setAttr(entry, "TDAElgi", "0");
					setAttr(entry, "TDSPercentage", "0");
					setAttr(entry, "PayRecAmount", "0");
					voucherRoot.appendChild(entry);
					childElements(sourceEntry).forEach(function (detail) {
						var accHead;
						if (String(detail.nodeName).toLowerCase() === "acchead") {
							accHead = doc.createElement("AccHead");
							setAttr(accHead, "No", attr(detail, "No"));
							setAttr(accHead, "CostCenter", attr(detail, "CostCenter"));
							setAttr(accHead, "Analytical", attr(detail, "Analytical"));
							setAttr(accHead, "Name", attr(detail, "Name"));
							setAttr(accHead, "Type", attr(detail, "Type"));
							setAttr(accHead, "TransFlag", "W");
							entry.appendChild(accHead);
						}
					});
					appendNarration(doc, entry, "");
					handledEntry = true;
				});
			}
		});
		if (!checkFinDate()) {
			return false;
		}
		return finishPost("XMLSave.asp?Mod=CNGJ&Name=Voucher Entry", doc, "VouCNGJGenerate.asp?hCallFrom=PI", disableNextButton);
	}

	function saveGeneralJournal(root) {
		var doc = xmlDocument("GJVoucher");
		var voucherRoot;
		var context = {};
		var counter = { value: 0 };
		var handledEntry = false;
		var taxRoot = null;
		if (!doc) {
			alert("Voucher XML data is not available.");
			return false;
		}
		clearXmlDocument(doc);
		voucherRoot = doc.createElement("voucher");
		doc.appendChild(voucherRoot);
		childElements(root).forEach(function (section) {
			var entries;
			if (String(section.nodeName).toLowerCase() === "header") {
				context = readHeaderContext(section, voucherRoot);
			} else if (String(section.nodeName).toLowerCase() === "details") {
				setAttr(voucherRoot, "VouDate", getVoucherDate());
				setAttr(voucherRoot, "BookAcchead", context.bookAccHead || "");
				setAttr(voucherRoot, "Approver", valueOf("selUserId"));
				entries = childElements(section, "Entry");
				if (entries.length && !handledEntry) {
					counter.value += 1;
					appendFirstGjEntry(doc, voucherRoot, context, entries[0], counter);
					counter.value += 1;
					appendSecondGjEntry(doc, voucherRoot, context, entries[0], counter, 1);
					handledEntry = true;
				}
			} else if (String(section.nodeName).toLowerCase() === "taxdetails") {
				taxRoot = section;
			}
		});
		if (taxRoot) {
			appendTaxGjEntries(doc, voucherRoot, taxRoot, context, counter);
		}
		if (!checkFinDate()) {
			return false;
		}
		return finishPost("XMLSave.asp?Mod=GJ&Name=Voucher Entry", doc, "VouGenerate.asp");
	}

	function saveXml() {
		var root = xmlRoot("TaxData");
		if (!root) {
			alert("Voucher XML data is not available.");
			return false;
		}
		if (isAmendmentPage() || isDirectNoteEntryPage()) {
			return saveCreditNote(root);
		}
		if (/^voucnpurinventry\.asp$/i.test(pageName())) {
			return savePurchaseCreditGeneralJournal(root);
		}
		return saveGeneralJournal(root);
	}

	function reTotalCr() {
		var root = xmlRoot("TaxData");
		var taxNodes = selectNodes(root, "//Tax");
		var taxValue = 0;
		var total;
		taxNodes.forEach(function (node) {
			var fieldName = "txtTaxValue" + attr(node, "CatCode") + attr(node, "TaxCode");
			var value = valueOf(fieldName, 0);
			setAttr(node, "TaxAmount", value);
			taxValue += toNumber(value);
		});
		total = formatNumber(toNumber(valueOf("txtTotal")) + taxValue, 2);
		setValue("txtInvValue", total);
		setValue("txtCrNoteValue", total);
		if (isAmendmentPage()) {
			(function (taxDetails) {
				if (taxDetails) {
					setAttr(taxDetails, 0, total);
				}
			}(selectNodes(root, "//TaxDetails")[0]));
		}
	}

	function applyGlReturn(value) {
		var root = xmlRoot("GLData");
		var entries = root ? childElements(root) : [];
		var parts;
		var valueRoot;
		if (!entries.length) {
			valueRoot = xmlRoot(value);
			entries = valueRoot ? childElements(valueRoot) : [];
		}
		if (entries.length) {
			setValue("hCrAccHead", attr(entries[0], "RetField0"));
			setHtml("spAccHead", attr(entries[0], "RetField5"));
			return true;
		}
		if (attr(root, "RetField0")) {
			setValue("hCrAccHead", attr(root, "RetField0"));
			setHtml("spAccHead", attr(root, "RetField5"));
			return true;
		}
		if (typeof value === "string") {
			parts = value.split(":");
			if (parts.length > 1) {
				setValue("hCrAccHead", parts[0]);
				setHtml("spAccHead", parts[5] || "");
				return true;
			}
		}
		return false;
	}

	function openGlDialog(url, args, callback) {
		var features = "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features, callback || function () {});
		} else {
			window.open(url, "_blank", "height=480,width=420,resizable=no,status=no,scrollbars=yes");
		}
	}

	function continueGlHead(url, args) {
		openGlDialog(url, args, function (returnedValue) {
			var root = xmlRoot(returnedValue);
			var action = trim(attr(root, "Action")).toUpperCase();
			var query = attr(root, "PassQuery");
			var parts;
			if (action === "CLOSE") {
				return;
			}
			if (action && action !== "DONE" && query) {
				continueGlHead(url.replace(/\?.*$/, "") + "?" + query, args);
				return;
			}
			if (typeof returnedValue === "string") {
				parts = returnedValue.split(":");
				if (parts.length <= 1 && trim(returnedValue) !== "") {
					continueGlHead(url.replace(/\?.*$/, "") + "?" + returnedValue, args);
					return;
				}
			}
			applyGlReturn(returnedValue);
		});
	}

	function showGlHead(orgId) {
		var bookNo = valueOf("hBookcode", valueOf("hBookCode"));
		var accHead = valueOf("hCrAccHead", "0");
		var isAmd = isAmendmentPage();
		var useLocalDialog = isAmd || /^vou(?:dnother|cnpur)inventry\.asp$/i.test(pageName());
		var url = (useLocalDialog ? "GLHeadSelection.asp" : "../../Common/GLHeadSelection.asp") +
			"?orgId=" + encodeURIComponent(orgId) +
			"&BookId=01&BookNo=" + encodeURIComponent(bookNo) +
			"&AccHead=" + encodeURIComponent(accHead || "0");
		continueGlHead(url, useLocalDialog ? "" : xmlObject("GLData"));
	}

	function accHead(source) {
		var select = source || field("SelAccountHd");
		if (select && select.value === "G") {
			showGlHead(valueOf("hOrgId", valueOf("hOrgID")));
		}
	}

	function setInvDate() {
		setVoucherDate(valueOf("hVouDate", valueOf("hSelVouDate")));
	}

	function initPage() {
		upgradeModern();
		if (isAmendmentPage()) {
			setRetVal(field("SelCrAgain"), "1");
			if (typeof window.popTax === "function") {
				window.popTax();
			}
			setInvDate();
			return;
		}
		if (trim(valueOf("hCallFromVoucher")) === "GJ") {
			(function (invoice) {
				setValue("txtNarration", "CR Note for " + invoice.no + " " + invoice.date);
			}(invoiceParts()));
		}
		if (field("SelCrAgain")) {
			setRetVal(field("SelCrAgain"), "0");
		}
		applyDateLimits();
	}

	function install() {
		upgradeModern();
		window.SaveXML = saveXml;
		window.EnbApp = enableApproval;
		window.SetRetVal = setRetVal;
		window.ResetTax = resetTax;
		window.ResetInvVal = resetCrVal;
		window.ResetCrVal = resetCrVal;
		window.CheckNoSer = checkNoSer;
		window.CheckFinDate = checkFinDate;
		window.AccHead = accHead;
		window.showGLHead = showGlHead;
		window.ReTotalCr = reTotalCr;
		window.SetInvDate = setInvDate;
		window.SetDate = setInvDate;
		window.Init = initPage;
		window.init = initPage;
		ready(initPage);
	}

	window.ITMSSalesReturnCreditNoteEntryCompat = {
		install: install
	};
}(window, document));
