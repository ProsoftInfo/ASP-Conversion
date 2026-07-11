(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value) {
		return toNumber(value).toFixed(2);
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || {};
	}

	function field(name) {
		var frm = form();
		var target = String(name).toLowerCase();
		var index;
		if (!frm || !frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === target) {
				return frm.elements[index];
			}
		}
		return document.getElementsByName(name)[0] || document.getElementById(name) || null;
	}

	function valueOf(name, fallback) {
		var item = field(name);
		return item && item.value != null ? item.value : fallback || "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function textOf(id) {
		var item = byId(id);
		return item ? trim(item.textContent || item.innerText || item.innerHTML || "") : "";
	}

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.innerHTML = value == null ? "" : String(value);
		}
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 && select.options[select.selectedIndex] ? select.options[select.selectedIndex].text : "";
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(nameOrObject) {
		if (typeof nameOrObject !== "string") {
			return nameOrObject;
		}
		ensureCompat();
		return window[nameOrObject] || document[nameOrObject] || byId(nameOrObject) || null;
	}

	function xmlDocument(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function xmlRoot(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function serializeXml(nameOrObject) {
		var doc = xmlDocument(nameOrObject);
		var root = xmlRoot(nameOrObject);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function clearDocument(name) {
		var doc = xmlDocument(name);
		while (doc && doc.firstChild) {
			doc.removeChild(doc.firstChild);
		}
		return doc;
	}

	function createNode(xmlName, nodeName) {
		var doc = xmlDocument(xmlName);
		return doc && doc.createElement ? doc.createElement(nodeName) : document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function attr(node, nameOrIndex) {
		var item;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			item = node.attributes.item(nameOrIndex);
			return item ? item.nodeValue : "";
		}
		return node.getAttribute(nameOrIndex) || "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function setAttrByIndex(node, index, value, fallbackName) {
		var item = node && node.attributes && node.attributes.item(index);
		if (item) {
			item.nodeValue = value == null ? "" : String(value);
		} else if (fallbackName) {
			setAttr(node, fallbackName, value);
		}
	}

	function importFor(parent, node) {
		if (node && parent && parent.ownerDocument && parent.ownerDocument.importNode && node.ownerDocument !== parent.ownerDocument) {
			return parent.ownerDocument.importNode(node, true);
		}
		return node;
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr.responseText || "";
	}

	function syncPost(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(body || null);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function rootFromDialog(value) {
		return xmlRoot(value) || value && value.nodeType === 9 && value.documentElement || value && value.nodeType === 1 && value || null;
	}

	function retFromGlRoot(root) {
		var entry = childElements(root, "Entry")[0] || childElements(root)[0];
		if (!entry) {
			return "";
		}
		return [0, 1, 2, 3, 4, 5, 6, 7].map(function (index) {
			return attr(entry, "RetField" + index);
		}).join(":");
	}

	function runGlDialog(url, callback) {
		openDialog(url, xmlObject("GLHeadData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
			var root = rootFromDialog(value);
			var action = root ? trim(attr(root, "Action")).toUpperCase() : "";
			var passQuery = root ? trim(attr(root, "PassQuery")) : "";
			var text = root ? "" : trim(value);
			if (root && action === "CLOSE") {
				return;
			}
			if (root && action && action !== "DONE" && passQuery) {
				runGlDialog("../../Common/GLHeadSelection.asp?" + passQuery, callback);
				return;
			}
			if (!root && text && text.split(":").length <= 1) {
				runGlDialog("../../Common/GLHeadSelection.asp?" + text, callback);
				return;
			}
			callback(root ? retFromGlRoot(root) : text);
		});
	}

	function getDateControl(name) {
		var control = field(name) || byId(name);
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		return window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate ? window.ITMSModernCompat.toDisplayDate(control.value) : control.value || "";
	}

	function setDateControl(name, value) {
		var control = field(name) || byId(name);
		if (!control || trim(value) === "") {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			control.value = window.ITMSModernCompat.toIsoDate(value);
		} else {
			control.value = value;
		}
	}

	function setDateLimits(name, minDate, maxDate) {
		var control = field(name) || byId(name);
		if (!control) {
			return;
		}
		if (typeof control.SetMinDate === "function") {
			control.SetMinDate(minDate);
		} else if (typeof control.setMinDate === "function") {
			control.setMinDate(minDate);
		}
		if (typeof control.SetMaxDate === "function") {
			control.SetMaxDate(maxDate);
		} else if (typeof control.setMaxDate === "function") {
			control.setMaxDate(maxDate);
		}
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var parsed;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			return new Date(Number(match[3]) < 100 ? Number(match[3]) + 2000 : Number(match[3]), Number(match[2]) - 1, Number(match[1]));
		}
		parsed = new Date(text);
		return isNaN(parsed.getTime()) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
	}

	function dateToDisplay(date) {
		function pad2(value) {
			return value < 10 ? "0" + value : String(value);
		}
		return pad2(date.getDate()) + "/" + pad2(date.getMonth() + 1) + "/" + date.getFullYear();
	}

	function setFinancialDate() {
		var fromYear = trim(valueOf("hFromYr"));
		var toYear = trim(valueOf("hToYr"));
		var minDate = fromYear ? "01/04/" + fromYear : "";
		var finEndText = toYear ? "31/03/" + toYear : "";
		var finEnd = parseDate(finEndText);
		var current = parseDate(valueOf("hCurrDate")) || new Date();
		var maxDate;
		if (!minDate || !finEnd) {
			return;
		}
		maxDate = finEnd.getTime() < current.getTime() ? finEndText : dateToDisplay(current);
		setDateLimits("ctlDate", minDate, maxDate);
		setDateControl("ctlDate", maxDate);
	}

	function initState() {
		var root = xmlRoot("VoucherData");
		window.VouRoot = root;
		window.EntryRoot = window.EntryRoot && window.EntryRoot.nodeType ? window.EntryRoot : xmlRoot("EntryData");
		if (!window.EntryRoot || String(window.EntryRoot.nodeName).toLowerCase() !== "entry") {
			window.clearXML();
		}
		if (!window.iEntryNo) {
			window.iEntryNo = nextEntryNo(root);
		}
		if (typeof window.bVouFlag !== "boolean") {
			window.bVouFlag = false;
		}
		if (typeof window.bSavFlag !== "boolean") {
			window.bSavFlag = false;
		}
		return { vouRoot: root, entryRoot: window.EntryRoot };
	}

	function nextEntryNo(root) {
		var max = 0;
		childElements(root, "Entry").forEach(function (entry) {
			max = Math.max(max, toNumber(attr(entry, "No")));
		});
		return max + 1 || 1;
	}

	function clearTable(tableName, startIndex) {
		var table = byId(tableName);
		var start = Number(startIndex) || 0;
		if (!table || !table.rows) {
			return null;
		}
		while (table.rows.length > start) {
			table.deleteRow(start);
		}
		return table;
	}

	function insertCell(row, html, className, align, valign, colspan) {
		if (typeof window.InsertCell === "function") {
			return window.InsertCell(row, 1, "", html, className || "ExcelDisplayCell", align || "left", valign || "top", 0, 0, colspan || 0, 0, "");
		}
		var cell = row.insertCell();
		cell.innerHTML = html == null ? "" : String(html);
		cell.className = className || "ExcelDisplayCell";
		if (align) {
			cell.align = align;
		}
		if (valign) {
			cell.vAlign = valign;
		}
		if (colspan) {
			cell.colSpan = colspan;
		}
		return cell;
	}

	function escapeHtml(value) {
		return String(value == null ? "" : value)
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;");
	}

	function entryNarration(entry) {
		var narration = "";
		childElements(entry, "Narration").forEach(function (node) {
			narration = node.textContent || "";
		});
		return narration;
	}

	function entryAccount(entry) {
		var account = "";
		childElements(entry, "AccHead").forEach(function (node) {
			if (attr(node, "Type") === "P") {
				account = attr(node, "Name") || attr(node, 3);
			} else {
				account = (attr(node, "No") || attr(node, 0)) + "-" + (attr(node, "Name") || attr(node, 3));
			}
		});
		return account;
	}

	function entryAccountName(entry) {
		var account = "";
		childElements(entry, "AccHead").forEach(function (node) {
			account = attr(node, "Name") || attr(node, 3);
		});
		return account;
	}

	function additionalDetails(entry) {
		var details = [];
		childElements(entry).forEach(function (header) {
			if (header.nodeName === "CostCenter" || header.nodeName === "Analytical") {
				childElements(header).forEach(function (node) {
					details.push(escapeHtml((attr(node, "ShortName") || attr(node, 2)) + "-" + (attr(node, "Ratio") || attr(node, 3)) + "% " + (attr(node, "Amount") || attr(node, 4))));
				});
			}
			if (header.nodeName === "PayRec") {
				childElements(header).forEach(function (node) {
					details.push(escapeHtml(attr(node, 1) + ":" + attr(node, 2) + "- " + attr(node, 5)));
				});
			}
		});
		return details.join("<br>");
	}

	function clearGeneratedEntryChildren(entryRoot) {
		childElements(entryRoot, "Narration").forEach(function (node) {
			entryRoot.removeChild(node);
		});
	}

	function addNarration(entryRoot, narrationText) {
		var narration = createNode("EntryData", "Narration");
		narration.textContent = narrationText == null ? valueOf("txtNarration") : narrationText;
		entryRoot.appendChild(narration);
	}

	function updateAnalysisAmounts(entryRoot) {
		childElements(entryRoot).forEach(function (header) {
			childElements(header).forEach(function (node) {
				var code = attr(node, "No") || attr(node, 0);
				var groupCode = attr(node, "GroupCode");
				var ratioField;
				var amountField;
				if (header.nodeName === "CostCenter") {
					ratioField = field("txtCCRatio" + code);
					amountField = field("txtCCAmount" + code);
				}
				if (header.nodeName === "Analytical") {
					ratioField = field("txtANALRatio" + code + "Z" + groupCode) || field("txtANALRatio" + code);
					amountField = field("txtANALAmount" + code + "Z" + groupCode) || field("txtANALAmount" + code);
				}
				if (header.nodeName === "PayRec") {
					amountField = field("txtDocAmount" + code);
				}
				if (ratioField) {
					setAttr(node, "Ratio", ratioField.value);
				}
				if (amountField) {
					if (header.nodeName === "PayRec") {
						setAttrByIndex(node, 5, amountField.value, "Amount");
					} else {
						setAttr(node, "Amount", amountField.value);
					}
				}
			});
		});
	}

	function applyTdsEligibility(value) {
		var eligible = trim(value) === "1";
		setValue("hTdsElgi", eligible ? "1" : "0");
		if (field("txtTdsAmount")) {
			field("txtTdsAmount").disabled = !eligible;
		}
		if (field("txtTdsper")) {
			field("txtTdsper").disabled = !eligible;
		}
	}

	function resetEntryFields() {
		var account = field("selAccountHead");
		if (account) {
			account.selectedIndex = 0;
		}
		setText("spAccHead", "");
		setValue("txtAmount", "0.00");
		setValue("txtNarration", field("txtNarration") ? field("txtNarration").defaultValue || "" : "");
		setValue("txtTdsAmount", "0.00");
		setValue("txtTdsper", "0.00");
		setValue("hTdsElgi", "0");
		setValue("hAction", "New");
		applyTdsEligibility("0");
		window.setADDDisplay(0);
		window.clearXML();
		window.bVouFlag = false;
	}

	function appendAccHeadFromData() {
		var state = initState();
		childElements(xmlRoot("AccHeadData")).forEach(function (node) {
			state.entryRoot.appendChild(importFor(state.entryRoot, node));
		});
	}

	function appendEntryInOrder(root, entry, entryNo) {
		var next = childElements(root, "Entry").filter(function (node) {
			return toNumber(attr(node, "No")) === toNumber(entryNo) + 1;
		})[0];
		if (next) {
			root.insertBefore(importFor(root, entry), next);
		} else {
			root.appendChild(importFor(root, entry));
		}
	}

	function finishPost(url, xml, action) {
		var xhr = syncPost(url, xml);
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		if (field("btnNext")) {
			field("btnNext").disabled = true;
		}
		if (field("btnAdd")) {
			field("btnAdd").disabled = true;
		}
		form().action = action;
		form().submit();
		return true;
	}

	function appendGjEntry(doc, root, no, crdr, payTo, amount, unitNo, unitName) {
		var entry = doc.createElement("Entry");
		setAttr(entry, "No", no);
		setAttr(entry, "CRDR", crdr);
		setAttr(entry, "Payto", payTo);
		setAttr(entry, "Amount", amount);
		setAttr(entry, "AccUnit", unitNo);
		setAttr(entry, "AccName", unitName);
		setAttr(entry, "TdsAmount", "0.00");
		setAttr(entry, "TDSElgi", "0");
		setAttr(entry, "TdsPercentage", "0");
		setAttr(entry, "PayRecAmount", "0");
		root.appendChild(entry);
		return entry;
	}

	function appendNarrationToGjEntry(doc, entry, narrationText) {
		var narration = doc.createElement("Narration");
		narration.textContent = narrationText || "";
		entry.appendChild(narration);
	}

	function buildGeneralJournalVoucher() {
		var sourceRoot = xmlRoot("VoucherData");
		var doc = clearDocument("GJVoucher");
		var root;
		var sourceEntry;
		var firstEntry;
		var secondEntry;
		var accHead;
		var recCount;
		var unitNo;
		var unitName;
		var amount;
		var partyParts;
		var payRecCount = [];
		var narrationText;
		if (!sourceRoot || !doc) {
			return null;
		}
		root = doc.createElement("voucher");
		doc.appendChild(root);
		unitNo = attr(sourceRoot, "UnitNo");
		unitName = attr(sourceRoot, "UnitName");
		setAttr(root, "UnitNo", unitNo);
		setAttr(root, "UnitName", unitName);
		setAttr(root, "BookNo", attr(sourceRoot, "BookNo"));
		setAttr(root, "BookName", attr(sourceRoot, "BookName"));
		setAttr(root, "CRDR", "");
		setAttr(root, "VouDate", attr(sourceRoot, "VouDate"));
		setAttr(root, "BookAcchead", "0");
		setAttr(root, "Approver", attr(sourceRoot, "Approver"));
		sourceEntry = childElements(sourceRoot, "Entry")[0];
		if (!sourceEntry) {
			return doc;
		}
		amount = attr(sourceEntry, "Amount");
		narrationText = entryNarration(sourceEntry) || valueOf("txtNarration");
		partyParts = String(attr(sourceRoot, "PartyCode")).split("?");
		payRecCount = trim(syncGet(
			"XMLGetPayRecCount.asp?orgID=" + encodeURIComponent(unitNo) +
			"&ParSubType=" + encodeURIComponent(partyParts[1] || "") +
			"&ParType=" + encodeURIComponent(partyParts[0] || "") +
			"&PartyCode=" + encodeURIComponent(partyParts[3] || "")
		)).split(":");
		firstEntry = appendGjEntry(doc, root, "1", "C", "0", amount, unitNo, unitName);
		accHead = doc.createElement("AccHead");
		setAttr(accHead, "No", attr(sourceRoot, "PartyCode"));
		setAttr(accHead, "Pay", payRecCount[0] || "");
		setAttr(accHead, "Rec", payRecCount[1] || "");
		setAttr(accHead, "Name", attr(sourceRoot, "PartyName"));
		setAttr(accHead, "Type", "P");
		setAttr(accHead, "Adv", payRecCount[2] || "");
		firstEntry.appendChild(accHead);
		recCount = doc.createElement("RecCount");
		setAttr(recCount, "Val", "1");
		firstEntry.appendChild(recCount);
		appendNarrationToGjEntry(doc, firstEntry, narrationText);
		secondEntry = appendGjEntry(doc, root, "2", "D", "", amount, unitNo, unitName);
		childElements(sourceEntry, "AccHead").forEach(function (sourceAcc) {
			var node = doc.createElement("AccHead");
			setAttr(node, "No", attr(sourceAcc, "No"));
			setAttr(node, "CostCenter", attr(sourceAcc, "CostCenter"));
			setAttr(node, "Analytical", attr(sourceAcc, "Analytical"));
			setAttr(node, "Name", attr(sourceAcc, "Name"));
			setAttr(node, "Type", "G");
			setAttr(node, "TransFlag", "A");
			secondEntry.appendChild(node);
		});
		return doc;
	}

	window.CheckAccHead = function (root, accHead) {
		return childElements(root, "Entry").some(function (entry) {
			return childElements(entry, "AccHead").some(function (node) {
				return trim(attr(node, "No")) === trim(accHead);
			});
		});
	};

	window.PopAccHead = function (select) {
		var state = initState();
		var parts;
		var node;
		var code;
		if (!select || select.selectedIndex <= 0) {
			return false;
		}
		if (select.value === "G") {
			return window.showGLHead(valueOf("hOrgId"));
		}
		parts = String(select.value || "").split("?");
		code = trim(parts[0]);
		applyTdsEligibility(parts[4]);
		if (window.CheckAccHead(state.vouRoot, code)) {
			alert("Account Head already Exisit in Voucher");
			select.selectedIndex = 0;
			window.bVouFlag = false;
			return false;
		}
		window.clearXML();
		state = initState();
		node = createNode("EntryData", "AccHead");
		setAttr(node, "No", code);
		setAttr(node, "CostCenter", trim(parts[1]));
		setAttr(node, "Analytical", trim(parts[2]));
		setAttr(node, "Name", selectedText(select));
		setAttr(node, "Type", "G");
		setAttr(node, "Group", "");
		state.entryRoot.appendChild(node);
		setText("spAccHead", selectedText(select));
		window.bVouFlag = true;
		if (field("txtNarration")) {
			field("txtNarration").focus();
		}
		if (typeof window.showCCAnal === "function") {
			window.showCCAnal(valueOf("hOrgId"), code, trim(parts[1]), trim(parts[2]));
		}
		return false;
	};

	window.showGLHead = function (orgId) {
		var base = "../../Common/GLHeadSelection.asp";
		var url = base + "?orgId=" + encodeURIComponent(orgId || valueOf("hOrgId")) + "&BookId=01&BookNo=" + encodeURIComponent(valueOf("hBookcode")) + "&AccHead=0";
		runGlDialog(url, function (ret) {
			var parts = String(ret || "").split(":");
			var accNode;
			var code;
			if (parts.length <= 1) {
				return;
			}
			applyTdsEligibility(parts[6]);
			clearChildren(xmlRoot("AccHeadData"));
			if (typeof window.GetGlHeadXml === "function") {
				window.GetGlHeadXml(ret);
			}
			accNode = childElements(xmlRoot("AccHeadData"), "AccHead")[0];
			code = accNode ? attr(accNode, "No") || attr(accNode, 0) : "";
			if (window.CheckAccHead(xmlRoot("VoucherData"), code)) {
				alert("Account Head already Exisit in Voucher");
				if (field("selAccountHead")) {
					field("selAccountHead").selectedIndex = 0;
				}
				clearChildren(xmlRoot("AccHeadData"));
				window.clearXML();
				window.bVouFlag = false;
				return;
			}
			window.clearXML();
			appendAccHeadFromData();
			accNode = childElements(window.EntryRoot, "AccHead")[0];
			if (accNode) {
				setText("spAccHead", attr(accNode, "Name") || attr(accNode, 3));
				window.bVouFlag = true;
				if (typeof window.showCCAnal === "function") {
					window.showCCAnal(valueOf("hOrgId"), attr(accNode, "No") || attr(accNode, 0), attr(accNode, "CostCenter") || attr(accNode, 1), attr(accNode, "Analytical") || attr(accNode, 2));
				}
			} else {
				if (field("selAccountHead")) {
					field("selAccountHead").selectedIndex = 0;
				}
				setText("spAccHead", "");
				window.setADDDisplay(0);
				window.bVouFlag = false;
			}
		});
		return false;
	};

	window.CheckVouStat = function () {
		var voucherDate = parseDate(getDateControl("ctlDate"));
		var currentDate = parseDate(valueOf("hCurrDate"));
		if (voucherDate && currentDate && voucherDate.getTime() > currentDate.getTime()) {
			alert("Voucher Date Should be Less than the System Date ");
			return false;
		}
		return true;
	};

	window.checkFileds = function () {
		if (trim(valueOf("txtNarration")) === "") {
			alert("Enter Narration");
			if (field("txtNarration")) {
				field("txtNarration").select();
			}
			return false;
		}
		if (typeof window.ValidateAmount === "function") {
			if (!window.ValidateAmount(valueOf("txtAmount"))) {
				if (field("txtAmount")) {
					field("txtAmount").select();
				}
				return false;
			}
		} else if (toNumber(valueOf("txtAmount")) <= 0) {
			alert("Amount should be >0 and < 9999999999.99");
			return false;
		}
		return true;
	};

	window.AddNew = function () {
		return window.AddEntry(trim(valueOf("hAction")) === "Edit" ? "U" : "A");
	};

	window.AddEntry = function (flag) {
		var state = initState();
		var entry = state.entryRoot;
		var entryNo;
		if (String(flag) === "S" && childElements(state.vouRoot, "Entry").length > 0 && formatNumber(valueOf("txtAmount")) === "0.00") {
			return window.CheckVouStat() ? window.SaveXML() : false;
		}
		if (!window.bVouFlag && !childElements(entry, "AccHead").length) {
			alert("Select an Account Head");
			return false;
		}
		if (!window.checkFileds()) {
			return false;
		}
		setAttr(state.vouRoot, "VouDate", getDateControl("ctlDate"));
		entryNo = String(flag) === "U" ? valueOf("hEntryNo") : nextEntryNo(state.vouRoot);
		setAttr(entry, "No", entryNo);
		setAttr(entry, "CRDR", "D");
		setAttr(entry, "Payto", textOf("spAccHead"));
		setAttr(entry, "Amount", valueOf("txtAmount"));
		setAttr(entry, "AccUnit", valueOf("hOrgId"));
		setAttr(entry, "AccName", valueOf("hOrgName"));
		setAttr(entry, "TdsAmount", valueOf("txtTdsAmount", "0.00"));
		setAttr(entry, "TDSElgi", valueOf("hTdsElgi", "0"));
		setAttr(entry, "TdsPercentage", valueOf("txtTdsper", "0.00"));
		clearGeneratedEntryChildren(entry);
		addNarration(entry);
		updateAnalysisAmounts(entry);
		if (String(flag) === "U") {
			appendEntryInOrder(state.vouRoot, entry, entryNo);
			setValue("hAction", "New");
		} else {
			state.vouRoot.appendChild(importFor(state.vouRoot, entry));
		}
		if (String(flag) === "S") {
			return window.CheckVouStat() ? window.SaveXML() : false;
		}
		window.DisplayVoucher();
		window.iEntryNo = nextEntryNo(state.vouRoot);
		setText("spEntryNo", String(window.iEntryNo));
		resetEntryFields();
		return false;
	};

	window.DisplayVoucher = function () {
		var root = initState().vouRoot;
		var table = clearTable("tblVoucher", 1);
		var display = byId("DisVoucher");
		var total = 0;
		var rowNo = 1;
		var row;
		if (!table) {
			return;
		}
		if (display) {
			display.style.height = String(60 + childElements(root, "Entry").length * 25) + "px";
			display.style.visibility = "visible";
		}
		childElements(root, "Entry").forEach(function (entry) {
			var amount = toNumber(attr(entry, "Amount"));
			setAttr(entry, "No", rowNo);
			total += amount;
			row = table.insertRow(table.rows.length);
			insertCell(row, rowNo, "ExcelSerial", "Center", "top");
			insertCell(row, '<img src="../../assets/images/iTMS%20icons/Deleteicon.gif" onclick="EditEntry(\'' + rowNo + '\',\'D\')">', "ExcelDisplayCell", "Center", "top");
			insertCell(row, '<a class="ExcelDisplayLink" href="javascript:EditEntry(\'' + rowNo + '\',\'E\')"><b>Edit</b></a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, escapeHtml(entryAccount(entry)), "ExcelDisplayCell", "left", "top");
			insertCell(row, escapeHtml(entryNarration(entry)), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(amount), "ExcelDisplayCell", "right", "top");
			insertCell(row, additionalDetails(entry), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(attr(entry, "TdsAmount")), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(attr(entry, "TdsPercentage")), "ExcelDisplayCell", "left", "top");
			rowNo += 1;
		});
		if (rowNo > 1) {
			row = table.insertRow(table.rows.length);
			insertCell(row, "<b>Total</b>", "ExcelDisplayCell", "right", "top", 5);
			insertCell(row, formatNumber(total), "ExcelDisplayCell", "right", "top");
			insertCell(row, "", "ExcelDisplayCell", "right", "top");
			insertCell(row, "", "ExcelDisplayCell", "right", "top");
			insertCell(row, "", "ExcelDisplayCell", "right", "top");
		}
		window.iEntryNo = rowNo;
		setText("spEntryNo", String(rowNo));
	};

	window.SaveXML = function () {
		var doc;
		if (trim(valueOf("hCallFrom")) === "CR") {
			if (typeof window.CheckApp === "function" && !window.CheckApp()) {
				return false;
			}
			return finishPost("XMLSave.asp?Mod=CN&Name=Voucher%20Entry", serializeXml("VoucherData"), "VouCNGenerate.asp");
		}
		doc = buildGeneralJournalVoucher();
		if (!doc) {
			alert("Voucher XML data is not available.");
			return false;
		}
		return finishPost("XMLSave.asp?Mod=GJ&Name=Voucher%20Entry", new XMLSerializer().serializeToString(doc), "VouGenerate.asp");
	};

	window.clearXML = function () {
		var entry = createNode("EntryData", "Entry");
		setAttr(entry, "No", window.iEntryNo || 1);
		setAttr(entry, "CRDR", "");
		setAttr(entry, "Payto", "");
		setAttr(entry, "Amount", "");
		setAttr(entry, "AccUnit", "");
		setAttr(entry, "AccName", "");
		setAttr(entry, "TdsAmount", "");
		setAttr(entry, "TDSElgi", "0");
		setAttr(entry, "TdsPercentage", "");
		window.EntryRoot = entry;
		return entry;
	};

	window.setAnalDisplay = function (display, flag) {
		var visible = Number(flag) !== 0;
		var item = byId(String(display) === "A" ? "DisAnal" : "DisCost");
		if (item) {
			item.style.height = visible ? "100px" : "1px";
			item.style.width = visible ? "280px" : "1px";
			item.style.visibility = visible ? "visible" : "hidden";
		}
	};

	window.setADDDisplay = function (flag) {
		var visible = Number(flag) !== 0;
		var additional = byId("Disaddtional");
		var ccanl = byId("DisCCANL");
		if (additional) {
			additional.style.height = visible ? "115px" : "1px";
			additional.style.visibility = visible ? "visible" : "hidden";
		}
		if (ccanl) {
			ccanl.style.height = visible ? "114px" : "1px";
			ccanl.style.visibility = visible ? "visible" : "hidden";
		}
	};

	window.CancelAction = function (page) {
		form().action = page;
		form().submit();
	};

	window.EditEntry = function (entryNo, editType) {
		var root = initState().vouRoot;
		var entry = childElements(root, "Entry").filter(function (node) {
			return String(attr(node, "No")) === String(entryNo);
		})[0];
		var accountSelect = field("selAccountHead");
		if (!entry) {
			return false;
		}
		setValue("hEntryNo", entryNo);
		setText("spEntryNo", entryNo);
		window.setADDDisplay(0);
		root.removeChild(entry);
		window.EntryRoot = entry;
		if (String(editType) === "D") {
			window.DelEntry();
			return false;
		}
		setValue("hAction", "Edit");
		setValue("txtAmount", attr(entry, "Amount"));
		setValue("txtTdsAmount", attr(entry, "TdsAmount") || "0.00");
		setValue("txtTdsper", attr(entry, "TdsPercentage") || "0.00");
		applyTdsEligibility(attr(entry, "TDSElgi"));
		setText("spAccHead", entryAccountName(entry));
		childElements(entry).forEach(function (node) {
			var index;
			if (node.nodeName === "AccHead" && accountSelect) {
				for (index = 0; index < accountSelect.options.length; index += 1) {
					if (String(accountSelect.options[index].value) === String(attr(node, "Type"))) {
						accountSelect.selectedIndex = index;
						break;
					}
				}
			}
			if (node.nodeName === "Narration") {
				setValue("txtNarration", node.textContent || "");
			}
			if (node.nodeName === "CostCenter" && typeof window.popCostCenter === "function") {
				window.setADDDisplay(1);
				window.popCostCenter(node);
			}
			if (node.nodeName === "Analytical" && typeof window.popAnalytical === "function") {
				window.setADDDisplay(1);
				window.popAnalytical(node);
			}
			if (node.nodeName === "PayRec" && typeof window.popPayRec === "function") {
				window.popPayRec(node);
			}
		});
		window.bVouFlag = true;
		window.DisplayVoucher();
		return false;
	};

	window.DelEntry = function () {
		window.clearXML();
		window.setADDDisplay(0);
		resetEntryFields();
		window.DisplayVoucher();
		window.bVouFlag = false;
		window.bSavFlag = true;
		return false;
	};

	window.showNarration = function () {
		var url = "NarrationSelection.asp?orgId=" + encodeURIComponent(valueOf("hOrgId")) + "&BookCode=" + encodeURIComponent("07?" + valueOf("hBookcode"));
		openDialog(url, "", "dialogHeight:300px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function (narration) {
			if (trim(narration) !== "") {
				setValue("txtNarration", narration);
			}
		});
	};

	window.InitVouCNOthersEntry = function () {
		ensureCompat();
		setFinancialDate();
		initState();
		applyTdsEligibility(valueOf("hTdsElgi", "0"));
		setText("spEntryNo", String(window.iEntryNo || 1));
		if (childElements(xmlRoot("VoucherData"), "Entry").length) {
			window.DisplayVoucher();
		}
	};
}(window, document));
