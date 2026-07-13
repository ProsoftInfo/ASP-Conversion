(function (window, document) {
	"use strict";

	var path = String(window.location.pathname || "").toLowerCase();

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form(name) {
		return document.forms[name || "formname"] || document.formname || document.forms[0] || null;
	}

	function field(name, formName) {
		var frm = form(formName);
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function fieldValue(name, fallback, formName) {
		var item = field(name, formName);
		return item ? item.value : fallback || "";
	}

	function setFieldValue(name, value, formName) {
		var item = field(name, formName);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.innerHTML = value == null ? "" : String(value);
		}
	}

	function asArray(collection) {
		return Array.prototype.slice.call(collection || []);
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		return asArray(node && node.childNodes).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name) || null;
	}

	function xmlDocument(name) {
		var object = typeof name === "string" ? xmlObject(name) : name;
		return object && object.XMLDocument || object && object._doc || object || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || null;
	}

	function dialogArgumentsValue() {
		var args = window.dialogArguments;
		var match;
		var id;
		if (!args && window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.dialogArgumentsRoot) {
			args = window.ITMSModalReturnCompat.dialogArgumentsRoot();
		}
		if (!args && window.opener && window.opener.__itmsDialogArgs) {
			match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
			id = match ? decodeURIComponent(match[1]) : "";
			if (id && Object.prototype.hasOwnProperty.call(window.opener.__itmsDialogArgs, id)) {
				args = window.opener.__itmsDialogArgs[id];
				window.dialogArguments = args;
			}
		}
		return args;
	}

	function createXmlElement(xmlName, nodeName) {
		var doc = xmlDocument(xmlName);
		if (doc && doc.createElement) {
			return doc.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function serializeXml(nodeOrDoc) {
		if (!nodeOrDoc) {
			return "";
		}
		if (typeof nodeOrDoc.xml === "string") {
			return nodeOrDoc.xml;
		}
		return new XMLSerializer().serializeToString(nodeOrDoc);
	}

	function loadXml(xmlName, xmlText) {
		var object = xmlObject(xmlName);
		var doc;
		if (object && object.loadXML) {
			object.loadXML(xmlText || "<Root/>");
			return object;
		}
		doc = new DOMParser().parseFromString(xmlText || "<Root/>", "text/xml");
		if (object) {
			object._doc = doc;
		}
		return object || doc;
	}

	function getXmlFromXhr(xhr) {
		if (xhr && xhr.responseXML && xhr.responseXML.documentElement) {
			return serializeXml(xhr.responseXML);
		}
		return xhr ? xhr.responseText || "" : "";
	}

	function createHttp() {
		return new XMLHttpRequest();
	}

	function syncGet(url) {
		var xhr = createHttp();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function syncPost(url, payload) {
		var xhr = createHttp();
		xhr.open("POST", url, false);
		xhr.send(payload);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
			return;
		}
		alert("Modern browser compatibility script is still loading. Please try again.");
	}

	function dialogId() {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
	}

	function notifyDialogValue(id, value) {
		if (!id || !window.opener) {
			return;
		}
		try {
			if (window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
				window.opener.ITMSModernCompat._receiveDialogValue(id, value);
				return;
			}
		} catch (ignoreDirectReturn) {}
		try {
			window.opener.postMessage({ type: "itms-dialog-return", id: id, value: value }, window.location.origin || "*");
		} catch (ignoreMessageReturn) {}
	}

	function returnDialogValue(value) {
		var id;
		window.returnValue = value;
		window.returnvalue = value;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
			return;
		}
		id = dialogId();
		notifyDialogValue(id, value);
	}

	function closeWithValue(value) {
		returnDialogValue(value);
		window.close();
	}

	function submitForm() {
		var frm = form();
		if (frm && typeof frm.submit === "function") {
			frm.submit();
		}
	}

	function dateControl(name) {
		ensureCompat();
		return field(name) || byId(name);
	}

	function setDateControl(name, value) {
		var control = dateControl(name);
		if (!control || !value) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else {
			control.value = window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate ? window.ITMSModernCompat.toIsoDate(value) : value;
		}
	}

	function getDateControl(name) {
		var control = dateControl(name);
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate) {
			return window.ITMSModernCompat.toDisplayDate(control.value);
		}
		return control.value || "";
	}

	function selectedRadioValue(name) {
		var radios = document.getElementsByName(name);
		for (var i = 0; i < radios.length; i += 1) {
			if (radios[i].checked) {
				return radios[i].value || "";
			}
		}
		return "";
	}

	function popupSize(type, fallbackProgram, fallbackHeight, fallbackWidth) {
		var value = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup(String(type)) : "";
		var parts = String(value || "").split(":");
		return {
			program: parts[0] || fallbackProgram,
			height: parts[1] || fallbackHeight || "500",
			width: parts[2] || fallbackWidth || "500"
		};
	}

	function dialogFeatures(height, width) {
		return "dialogHeight:" + height + "px;dialogWidth:" + width + "px;center:Yes;help:No;resizable:No;status:No";
	}

	function pageBase(url) {
		return String(url || "").split("?")[0];
	}

	function continueSelectionDialog(url, args, features, callback) {
		openDialog(url, args, features, function (value) {
			var outRoot = xmlRoot(value) || value;
			var action = trim(attr(outRoot, "Action")).toUpperCase();
			var query;
			if (!outRoot || action === "CLOSE") {
				return;
			}
			if (action && action !== "DONE") {
				query = trim(attr(outRoot, "PassQuery"));
				continueSelectionDialog(pageBase(url) + (query ? "?" + query : ""), args, features, callback);
				return;
			}
			callback(outRoot);
		});
	}

	function checkedReminder() {
		var count = parseInt(fieldValue("hCnt", "0"), 10) || 0;
		var selected = [];
		var item;
		for (var i = 1; i <= count; i += 1) {
			item = field("chkReminderZ" + i);
			if (item && item.checked) {
				selected.push({
					value: item.value,
					passValue: fieldValue("hPartyDetZ" + i)
				});
			}
		}
		return selected;
	}

	function addCell(row, className, align, text) {
		var cell = row.insertCell(-1);
		if (className) {
			cell.className = className;
		}
		if (align) {
			cell.align = align;
		}
		if (text != null) {
			cell.innerHTML = String(text);
		}
		return cell;
	}

	function appendInput(cell, type, name, value, className, size, textAlign, checked) {
		var input = document.createElement("input");
		input.type = type || "text";
		input.name = name || "";
		input.value = value == null ? "" : String(value);
		if (className) {
			input.className = className;
		}
		if (size) {
			input.size = size;
		}
		if (textAlign) {
			input.style.textAlign = textAlign;
		}
		if (type && type.toLowerCase() === "checkbox") {
			input.checked = !!checked;
		} else {
			input.readOnly = true;
		}
		cell.appendChild(input);
		return input;
	}

	function clearTableRows(tableId, keepRows) {
		var table = byId(tableId);
		if (!table || !table.rows) {
			return;
		}
		while (table.rows.length > keepRows) {
			table.deleteRow(keepRows);
		}
	}

	function numberValue(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatAmount(value) {
		return numberValue(value).toFixed(2);
	}

	function sameNodePair(node, parts) {
		return trim(attr(node, "RetField1")) === trim(parts[1]) && trim(attr(node, "RetField2")) === trim(parts[2]);
	}

	function removeMatchingEntries(root, parts) {
		childElements(root, "Entry").forEach(function (entry) {
			if (sameNodePair(entry, parts)) {
				root.removeChild(entry);
			}
		});
	}

	function makeEntry(xmlName, values) {
		var node = createXmlElement(xmlName, "Entry");
		values.forEach(function (value, index) {
			setAttr(node, "RetField" + index, value);
		});
		return node;
	}

	function currentPageIsPaymentRequest() {
		return path.indexOf("paymentrequests.asp") !== -1;
	}

	function currentPageIsPaymentReminder() {
		return path.indexOf("paymentreminders.asp") !== -1;
	}

	function currentPageIsOverdueReminder() {
		return path.indexOf("overduereminders.asp") !== -1;
	}

	function currentPageIsPartyOutstanding() {
		return path.indexOf("partyoutstanding.asp") !== -1 && path.indexOf("partyoutstandingbreakup.asp") === -1;
	}

	function currentPageIsBreakup() {
		return path.indexOf("partyoutstandingbreakup.asp") !== -1;
	}

	function selectedPaymentRequests() {
		var count = parseInt(fieldValue("hCnt", "0"), 10) || 0;
		var selected = [];
		var check;
		for (var i = 1; i <= count; i += 1) {
			check = field("chkbox" + i);
			if (check && check.checked) {
				selected.push(String(check.value || "").split(":"));
			}
		}
		return selected;
	}

	function displayOutstandingData() {
		var root = xmlRoot("OutStandingData");
		var rowNo = 0;
		clearTableRows("RecTab", 2);
		if (!root) {
			setFieldValue("hCount", "");
			return;
		}
		childElements(root, "Party").forEach(function (party) {
			var partyCode = attr(party, "CODE");
			var partyType = attr(party, "TYPE");
			var partySubType = attr(party, "SUBTYPE");
			var partyName = attr(party, "NAME");
			childElements(party, "DETAILS").forEach(function (detail) {
				var row = byId("RecTab").insertRow(-1);
				var passValue = [partyCode, partyType, partySubType, attr(detail, "INVOICENO")].join(":");
				rowNo += 1;
				addCell(row, "ExcelSerial", "center", rowNo);
				appendInput(addCell(row, "ExcelDisplayCell", "center"), "checkbox", "ChkBox" + rowNo, passValue, "Formelem");
				appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "sPartyName", partyName, "Formelem", null, "Left");
				appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "InvAmt", formatAmount(attr(detail, "BALANCE")), "FormelemRead", 15, "Right");
				appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "nTDSAmt", "", "FormelemRead", 15, "Right");
				appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "nAdvance", "", "FormelemRead", 15, "Right");
				appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "nCNode", "", "FormelemRead", 15, "Right");
				appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "nFreight", "", "FormelemRead", 15, "Right");
				appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "nAmtPaid", formatAmount(attr(detail, "AMOUNTPAIDTILLDATE")), "FormelemRead", 15, "Right");
				appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "nAmtOut", formatAmount(attr(detail, "BALANCE")), "FormelemRead", 15, "Right");
			});
		});
		setFieldValue("hCount", rowNo || "");
	}

	function selectedOutstandingValues() {
		var count = parseInt(fieldValue("hCount", "0"), 10) || 0;
		var checked = [];
		var all = [];
		var item;
		for (var i = 1; i <= count; i += 1) {
			item = field("ChkBox" + i);
			if (item) {
				all.push(item.value);
				if (item.checked) {
					checked.push(item.value);
				}
			}
		}
		return checked.length ? checked : all;
	}

	function markSelectedReminderDetails(root) {
		var selected = trim(fieldValue("hSelNode"));
		var selectedMap = {};
		if (!selected || !root) {
			return;
		}
		selected.split(",").forEach(function (entry) {
			var parts = entry.split(":");
			if (parts.length >= 2) {
				selectedMap[trim(parts[0]) + ":" + trim(parts[1])] = true;
			}
		});
		childElements(root).forEach(function (party) {
			childElements(party, "DETAILS").forEach(function (detail) {
				if (selectedMap[trim(attr(detail, "SNO")) + ":" + trim(attr(detail, "INVOICENO"))]) {
					setAttr(detail, "SELECT", "Y");
				}
			});
		});
	}

	function buildPassValue(values) {
		var partyCodes = [];
		var partyTypes = [];
		var partySubTypes = [];
		var invoices = [];
		values.forEach(function (value) {
			var parts = String(value || "").split(":");
			partyCodes.push(parts[0] || "");
			partyTypes.push(parts[1] || "");
			partySubTypes.push(parts[2] || "");
			invoices.push(parts[3] || "");
		});
		return {
			partyType: trim(partyTypes.join(",")),
			value: [partyCodes.join(","), partyTypes.join(","), partySubTypes.join(","), invoices.join(",")].join(":")
		};
	}

	function checkOutstandingSubmit() {
		var count = trim(fieldValue("hCount"));
		var values;
		var pass;
		var filename;
		var xhr;
		var root;
		var responseParts;
		if (!count) {
			alert("Select Party Outstanding Break up");
			return;
		}
		values = selectedOutstandingValues();
		pass = buildPassValue(values);
		filename = pass.partyType === "CR" ? "XMLGenReminderPayables.asp" : "XMLGenReminder.asp";
		xhr = syncGet(filename + "?sData=" + encodeURIComponent(pass.value));
		if (trim(getXmlFromXhr(xhr))) {
			loadXml("GenReminder", getXmlFromXhr(xhr));
		}
		root = xmlRoot("GenReminder");
		if (root && childElements(root).length) {
			setAttr(root, "SENDBY", selectedRadioValue("radSendBy"));
			setAttr(root, "NAME", trim(fieldValue("txtCouComName")));
			setAttr(root, "ID", trim(fieldValue("txtCouTransID")));
			setAttr(root, "ADDRESS", trim(fieldValue("txtCouComAddress")));
		}
		markSelectedReminderDetails(root);
		xhr = syncPost("GenReminderInsert.asp", serializeXml(xmlDocument("GenReminder")));
		if (!trim(xhr.responseText)) {
			return;
		}
		responseParts = xhr.responseText.split("@");
		if (trim(responseParts[0])) {
			alert(responseParts[0]);
			return;
		}
		openDialog(
			"PartyOutstandingPrevReminder.asp?" + (responseParts[1] || "") + "&PassValue=" + encodeURIComponent(pass.value),
			xmlObject("OutStandingData"),
			dialogFeatures(450, 700),
			function () {
				window.location.href = pass.partyType === "CR" ? "PAYMENTREMINDERS.ASP" : "OverdueReminders.asp";
			}
		);
	}

	function selectedBreakupInvoices() {
		var selected = {};
		var root = xmlRoot(dialogArgumentsValue());
		var partyCode = trim(fieldValue("hPartycode"));
		childElements(root, "Party").forEach(function (party) {
			if (trim(attr(party, "CODE")) !== partyCode) {
				return;
			}
			childElements(party, "DETAILS").forEach(function (detail) {
				trim(attr(detail, "INVOICENO")).split(",").forEach(function (invoiceNo) {
					if (trim(invoiceNo)) {
						selected[trim(invoiceNo)] = true;
					}
				});
			});
		});
		return selected;
	}

	function displayBreakupData() {
		var root = xmlRoot("ReceivableData");
		var table = byId("RecTab");
		var selectedInvoices = selectedBreakupInvoices();
		if (!root || !table) {
			return;
		}
		clearTableRows("RecTab", 2);
		childElements(root).forEach(function (node) {
			if (String(node.nodeName).toUpperCase() === "PARTY") {
				childElements(node, "DETAILS").forEach(function (detail) {
					var row = table.insertRow(-1);
					var serialNo = attr(detail, "SNO");
					var invoiceNo = attr(detail, "INVOICENO");
					addCell(row, "ExcelSerial", "center", serialNo);
					appendInput(addCell(row, "ExcelDisplayCell", "center"), "checkbox", "ChkBox" + serialNo, serialNo + ":" + invoiceNo, "Formelem", null, null, !!selectedInvoices[trim(invoiceNo)]);
					appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "DocType", attr(detail, "DOCTYPE"), "Formelem", 6, "Center");
					appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "InvNo", invoiceNo, "Formelem", 5, "Center");
					appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "Date", attr(detail, "DATE"), "FormelemRead", 11, "Center");
					appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "AccOn", attr(detail, "ACCOUNTEDON"), "FormelemRead", 11, "Center");
					appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "InvAmt", attr(detail, "AMOUNT"), "FormelemRead", 15, "Right");
					appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "AmtPaidTill", attr(detail, "AMOUNTPAIDTILLDATE"), "FormelemRead", 15, "Right");
					appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "Balance", attr(detail, "BALANCE"), "FormelemRead", 15, "Right");
					appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "NoOfDaysOut", attr(detail, "NOOFDAYSOUT"), "FormelemRead", 7, "Center");
					appendInput(addCell(row, "ExcelDisplayCell", "Left"), "text", "NoOfDaysOver", attr(detail, "NOOFDAYSOVER"), "FormelemRead", 7, "Center");
					setFieldValue("hCnt", serialNo);
				});
			}
			if (String(node.nodeName).toUpperCase() === "TOTAL") {
				var totalRow = table.insertRow(-1);
				var cell = addCell(totalRow, "ExcelHeaderCell", "Right", "Total Outstanding Transaction Amount");
				cell.colSpan = 6;
				appendInput(addCell(totalRow, "ExcelDisplayCell", "Left"), "text", "InvAmt", attr(node, "AMOUNT"), "FormelemRead", 15, "Right");
				appendInput(addCell(totalRow, "ExcelDisplayCell", "Left"), "text", "AmtPaidTill", attr(node, "PAID"), "FormelemRead", 15, "Right");
				appendInput(addCell(totalRow, "ExcelDisplayCell", "Left"), "text", "Balance", attr(node, "RECEIVABLE"), "FormelemRead", 15, "Right");
				appendInput(addCell(totalRow, "ExcelDisplayCell", "Left"), "text", "NoOfDaysOut", "", "FormelemRead", 7, "Center");
				appendInput(addCell(totalRow, "ExcelDisplayCell", "Left"), "text", "NoOfDaysOver", "", "FormelemRead", 7, "Center");
			}
		});
	}

	function checkedBreakupRows() {
		return asArray(document.querySelectorAll('input[type="checkbox"][name^="ChkBox"]')).filter(function (item) {
			return item.checked;
		});
	}

	function addToReminder() {
		var root = xmlRoot("ReceivableData");
		var outRoot = xmlRoot("OutData");
		var temp = fieldValue("hTempVal").split(":");
		var partyNode;
		var selectedNode = [];
		var invoices = [];
		var amount = 0;
		var paid = 0;
		var balance = 0;
		if (!root || !outRoot) {
			closeWithValue(xmlRoot("OutData"));
			return;
		}
		while (outRoot.firstChild) {
			outRoot.removeChild(outRoot.firstChild);
		}
		partyNode = createXmlElement("OutData", "Party");
		setAttr(partyNode, "CODE", temp[0] || "");
		setAttr(partyNode, "TYPE", temp[1] || "");
		setAttr(partyNode, "SUBTYPE", temp[2] || "");
		setAttr(partyNode, "NAME", temp[3] || "");
		outRoot.appendChild(partyNode);
		checkedBreakupRows().forEach(function (check) {
			var parts = String(check.value || "").split(":");
			childElements(root, "Party").forEach(function (party) {
				childElements(party, "DETAILS").forEach(function (detail) {
					if (trim(attr(detail, "SNO")) === trim(parts[0]) && trim(attr(detail, "INVOICENO")) === trim(parts[1])) {
						selectedNode.push(check.value);
						setAttr(detail, "SELECTION", "Y");
						invoices.push(attr(detail, "INVOICENO"));
						amount += numberValue(attr(detail, "AMOUNT"));
						paid += numberValue(attr(detail, "AMOUNTPAIDTILLDATE"));
						balance += numberValue(attr(detail, "BALANCE"));
					}
				});
			});
		});
		if (selectedNode.length) {
			var detailNode = createXmlElement("OutData", "DETAILS");
			setAttr(detailNode, "INVOICENO", invoices.join(","));
			setAttr(detailNode, "AMOUNT", amount);
			setAttr(detailNode, "AMOUNTPAIDTILLDATE", paid);
			setAttr(detailNode, "BALANCE", balance);
			setAttr(detailNode, "Selected", selectedNode.join(","));
			partyNode.appendChild(detailNode);
		}
		closeWithValue(outRoot);
	}

	window.SetDate = function () {
		if (field("hFromDate") && field("hToDate")) {
			setDateControl("ctlVouFromDate", fieldValue("hFromDate"));
			setDateControl("ctlVouToDate", fieldValue("hToDate"));
		}
		if (field("hTillDate")) {
			setDateControl("ctlTillDate", fieldValue("hTillDate"));
		}
	};

	window.AssignPage = function (pageNo) {
		setFieldValue("hPage", pageNo);
		submitForm();
	};

	window.PaginateAcc = function (pageNo) {
		setFieldValue("hPageSelection", pageNo);
		submitForm();
	};

	window.ChkReset = function () {
		var frm = form();
		if (frm && typeof frm.reset === "function") {
			frm.reset();
		}
		setText("PartyName", "");
		window.SetDate();
	};

	window.Validate = function () {
		if (field("ctlVouFromDate") || byId("ctlVouFromDate")) {
			setFieldValue("hFromDate", getDateControl("ctlVouFromDate"));
		}
		if (field("ctlVouToDate") || byId("ctlVouToDate")) {
			setFieldValue("hToDate", getDateControl("ctlVouToDate"));
		}
		if (field("ctlTillDate") || byId("ctlTillDate")) {
			setFieldValue("hTillDate", getDateControl("ctlTillDate"));
		}
		if (field("selReqType")) {
			setFieldValue("hReqType", fieldValue("selReqType"));
		}
		if (field("chkCreated") || field("chkApproved")) {
			var created = field("chkCreated");
			var approved = field("chkApproved");
			var status = "";
			if (created && created.checked) {
				status = created.value;
			}
			if (approved && approved.checked) {
				status = approved.value;
			}
			if (created && approved && created.checked && approved.checked) {
				status = created.value + "','" + approved.value;
			}
			setFieldValue("hStatus", status);
		}
		if (document.getElementsByName("RadStatus").length) {
			setFieldValue("hsentBy", selectedRadioValue("RadStatus"));
		}
		submitForm();
	};

	window.Create = function () {
		openDialog("PaymtReqTypeSelection.asp", "", dialogFeatures(220, 350), function (value) {
			var requestType = trim(value);
			if (!requestType) {
				return;
			}
			setFieldValue("hReqTypeS", requestType);
			if (requestType === "H" || requestType === "L") {
				form().action = "PmtReqLoanEntry.asp";
			} else if (requestType === "B") {
				form().action = "PmtReqBlankChqEntry.asp";
			} else if (requestType === "A" || requestType === "O") {
				form().action = "PmtReqChequeEntry.asp";
			}
			submitForm();
		});
	};

	window.CheckSubmit = function (callFrom) {
		var selected = selectedPaymentRequests();
		var requestNos = [];
		var lastStatus = "";
		var lastRequestFrom = "";
		var blocked = false;
		if (!currentPageIsPaymentRequest()) {
			return;
		}
		selected.forEach(function (parts) {
			if (blocked) {
				return;
			}
			requestNos.push(parts[0] || "");
			lastStatus = parts[1] || "";
			lastRequestFrom = parts[2] || "";
			if (trim(lastStatus) === "010203" && trim(callFrom) !== "E" && trim(callFrom) !== "D" && trim(callFrom) !== "G") {
				alert("Already This Payment is approved");
				blocked = true;
			}
		});
		if (blocked) {
			return;
		}
		if (!requestNos.length) {
			alert("Select any one payment");
			return;
		}
		if (trim(callFrom) === "E" && trim(lastRequestFrom) !== "1") {
			alert("Selected Payment was created from other Module.Select any other payment");
			return;
		}
		if (callFrom === "A") {
			form().action = "PaymtReqApprove.asp?RequestNo=" + encodeURIComponent(requestNos.join(","));
		} else if (callFrom === "D") {
			form().action = "PaymtReqDeletion.asp?RequestNo=" + encodeURIComponent(requestNos.join(","));
		} else if (callFrom === "G") {
			form().action = "PmtGenReqularChqEntryforParty.asp?RequestNo=" + encodeURIComponent(requestNos.join(","));
		}
		submitForm();
	};

	window.DelSubmit = function () {
		var selected = checkedReminder();
		var xhr;
		if (selected.length !== 1) {
			alert("Select Any One Record to Delete");
			return;
		}
		if (!confirm("Do you want to Delete the Reminder Permanently")) {
			return;
		}
		xhr = syncGet("PayRecReminderDelete.asp?RemNo=" + encodeURIComponent(selected[0].value));
		if (!trim(xhr.responseText)) {
			alert("Reminder Deleted Successfully");
			submitForm();
		}
	};

	window.ViewRem = function () {
		var selected = checkedReminder();
		if (selected.length !== 1) {
			alert("Select Any One Record to Print");
			return;
		}
		openDialog(
			"PartyOutstandingPrevReminder.asp?PassValue=" + encodeURIComponent(selected[0].passValue) + "&RemNo=" + encodeURIComponent(selected[0].value) + "&CallFrom=List",
			xmlObject("OutStandingData"),
			dialogFeatures(450, 700)
		);
	};

	window.SelParty = function () {
		var size = popupSize("2", "PartySelection.asp", "500", "500");
		var args = xmlObject("PartyData");
		var unitId = fieldValue("hUnitNo");
		var wantsSubType = !!field("hPartySubType");
		var url = "../../Common/" + size.program + "?orgid=" + encodeURIComponent(unitId) + (wantsSubType ? "" : "&hSelectMode=M");
		setFieldValue("hPartyCode", "");
		if (wantsSubType) {
			setFieldValue("hPartySubType", "");
		}
		setText("PartyName", "");
		continueSelectionDialog(url, args, dialogFeatures(size.height, size.width), function (outRoot) {
			var codes = [];
			var names = [];
			var subTypes = [];
			childElements(outRoot, "Entry").forEach(function (entry) {
				codes.push(attr(entry, "RetField1"));
				names.push(attr(entry, "RetField0"));
				subTypes.push(attr(entry, "RetField4"));
			});
			if (names.length) {
				setFieldValue("hPartyCode", codes.join(","));
				if (wantsSubType) {
					setFieldValue("hPartySubType", subTypes.join(","));
				}
				setText("PartyName", names.join(","));
			}
		});
	};

	window.CreateRem = function () {
		if (currentPageIsPaymentReminder()) {
			form().action = "PartyOutstanding.asp?PartyType=CR";
		} else if (currentPageIsOverdueReminder()) {
			form().action = "PartyOutstanding.asp?PartyType=DR";
		} else {
			form().action = "PartyOutstanding.asp";
		}
		submitForm();
	};

	window.ShowCCDetails = function (obj) {
		var passValue = [
			fieldValue("hUnitNo"),
			fieldValue("hUnitName"),
			fieldValue("hTillDate"),
			obj && obj.name || ""
		].join("|");
		openDialog("PartyOutstandingBreakup.asp?Value=" + encodeURIComponent(passValue), xmlObject("OutStandingData"), dialogFeatures(450, 700), function (value) {
			var outRoot = xmlRoot(value) || value;
			var root = xmlRoot("OutStandingData");
			if (!outRoot || !root) {
				return;
			}
			childElements(outRoot, "Party").forEach(function (party) {
				var partyNode = createXmlElement("OutStandingData", "Party");
				setAttr(partyNode, "CODE", attr(party, "CODE"));
				setAttr(partyNode, "TYPE", attr(party, "TYPE"));
				setAttr(partyNode, "SUBTYPE", attr(party, "SUBTYPE"));
				setAttr(partyNode, "NAME", attr(party, "NAME"));
				childElements(party, "DETAILS").forEach(function (detail) {
					var detailNode = createXmlElement("OutStandingData", "DETAILS");
					setAttr(detailNode, "INVOICENO", attr(detail, "INVOICENO"));
					setAttr(detailNode, "AMOUNT", attr(detail, "AMOUNT"));
					setAttr(detailNode, "AMOUNTPAIDTILLDATE", attr(detail, "AMOUNTPAIDTILLDATE"));
					setAttr(detailNode, "BALANCE", attr(detail, "BALANCE"));
					setFieldValue("hSelNode", attr(detail, "Selected"));
					partyNode.appendChild(detailNode);
				});
				root.appendChild(partyNode);
			});
			displayOutstandingData();
		});
	};

	window.ClearTable = function () {
		clearTableRows("RecTab", currentPageIsBreakup() ? 2 : 2);
	};

	window.DisplayData = currentPageIsBreakup() ? displayBreakupData : displayOutstandingData;
	window.CheckSumbit = checkOutstandingSubmit;

	window.GetXML = function (value, partyType) {
		var filename = trim(partyType) === "CR" ? "XMLGetOutstandingPayables.asp" : "XMLGetOutstandingReceivables.asp";
		var xhr = syncGet(filename + "?Value=" + encodeURIComponent(value || ""));
		if (trim(getXmlFromXhr(xhr))) {
			loadXml("ReceivableData", getXmlFromXhr(xhr));
		}
		displayBreakupData();
	};

	window.CloseWindow = function () {
		window.close();
	};

	window.InvoicePopUp = function (rcvbleNo, billDate, paidAmount, balanceAmount, transType) {
		var value = [rcvbleNo, trim(billDate), paidAmount, balanceAmount, transType].join(":");
		openDialog("InvoicePopUp.asp?sTemp=" + encodeURIComponent(value), "", dialogFeatures(150, 450));
	};

	window.ShowDetails = function (transNo, days) {
		if (numberValue(days) > 0) {
			openDialog("ReceivablesPenalty.asp?TransNo=" + encodeURIComponent(transNo + "-" + days), "A", dialogFeatures(300, 700));
		}
	};

	window.AddToReminder = addToReminder;

	if (currentPageIsPartyOutstanding() || currentPageIsBreakup()) {
		ensureCompat();
	}
}(window, document));
