(function (window, document) {
	"use strict";

	function isEditMode() {
		return window.PurchaseVoucherEntryMode === "edit";
	}

	function isAmdDetailsMode() {
		return window.PurchaseVoucherEntryVariant === "amd-details";
	}

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
		return document.getElementsByName(name)[0] || null;
	}

	function fields(name) {
		var item = field(name);
		var result = [];
		var frm = form();
		var target = String(name).toLowerCase();
		var index;
		if (item && item.length != null && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		if (item) {
			return [item];
		}
		if (!frm || !frm.elements) {
			return result;
		}
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === target) {
				result.push(frm.elements[index]);
			}
		}
		return result;
	}

	function valueOf(name, fallback) {
		var item = field(name);
		return item ? item.value : fallback || "";
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

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.innerHTML = value == null ? "" : String(value);
		}
	}

	function textOf(id) {
		var item = byId(id);
		return item ? trim(item.textContent || item.innerText || item.innerHTML || "") : "";
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
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

	function loadXml(name, xmlText) {
		var object = xmlObject(name);
		var text = trim(xmlText) || "<Root/>";
		if (object && typeof object.loadXML === "function") {
			object.loadXML(text);
			return object;
		}
		if (object) {
			object._doc = new DOMParser().parseFromString(text, "text/xml");
		}
		return object;
	}

	function createNode(xmlName, nodeName) {
		var object = xmlObject(xmlName);
		if (object && typeof object.createElement === "function") {
			return object.createElement(nodeName);
		}
		if (object && object.XMLDocument) {
			return object.XMLDocument.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
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

	function selectNodes(context, expression) {
		var doc;
		var found;
		var result = [];
		var i;
		if (!context) {
			return result;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return result;
		}
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (i = 0; i < found.snapshotLength; i += 1) {
			result.push(found.snapshotItem(i));
		}
		return result;
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function syncPost(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(body);
		return xhr;
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
		if (!control) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else {
			control.value = window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate ? window.ITMSModernCompat.toIsoDate(value) : value || "";
		}
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var months = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
		var parsed;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{1,2})[\/-]([A-Za-z]{3})[\/-](\d{4})$/);
		if (match && months[match[2].toLowerCase()] != null) {
			return new Date(Number(match[3]), months[match[2].toLowerCase()], Number(match[1]));
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			return new Date(Number(match[3]) < 100 ? Number(match[3]) + 2000 : Number(match[3]), Number(match[2]) - 1, Number(match[1]));
		}
		parsed = new Date(text);
		return isNaN(parsed.getTime()) ? null : parsed;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
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

	function rootFromDialog(value) {
		return xmlRoot(value) || value && value.nodeType === 9 && value.documentElement || value && value.nodeType === 1 && value || null;
	}

	function runXmlDialog(url, nextPrefix, args, features, done) {
		openDialog(url, args || "", features, function (value) {
			var root = rootFromDialog(value);
			var action = trim(attr(root, "Action")).toUpperCase();
			var query = trim(attr(root, "PassQuery"));
			if (!root || action === "CLOSE") {
				return;
			}
			if (action && action !== "DONE" && query) {
				runXmlDialog(nextPrefix + query, nextPrefix, args, features, done);
				return;
			}
			done(root);
		});
	}

	function importFor(parent, node) {
		if (!parent || !node) {
			return node;
		}
		if (node.ownerDocument !== parent.ownerDocument && parent.ownerDocument.importNode) {
			return parent.ownerDocument.importNode(node, true);
		}
		return node;
	}

	function clearTable(tableName, startIndex, keepCount) {
		var table = byId(tableName);
		var start = Number(startIndex) || 0;
		var keep = Number(keepCount) || 0;
		if (!table || !table.rows) {
			return null;
		}
		while (table.rows.length > start + keep) {
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

	function selectByValue(select, value) {
		var wanted = trim(value).replace(/\|/g, "?");
		var optionValue;
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			optionValue = trim(select.options[i].value).replace(/\|/g, "?");
			if (optionValue === wanted) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function setEntryButtons(editing) {
		var add = field("btnAdd");
		var next = field("btnNext");
		var update = field("btnUpdate");
		var del = field("btnDel");
		if (add) {
			add.disabled = !!editing;
		}
		if (next) {
			next.disabled = !!editing;
		}
		if (update) {
			update.disabled = !editing;
		}
		if (del) {
			del.disabled = !editing;
		}
	}

	function selectedRadioValue(name, fallback) {
		var radios = fields(name);
		for (var i = 0; i < radios.length; i += 1) {
			if (radios[i].checked) {
				return radios[i].value;
			}
		}
		return fallback || "";
	}

	function nextEntryNo(root) {
		var maxNo = 0;
		childElements(root, "Entry").forEach(function (entry) {
			maxNo = Math.max(maxNo, Number(attr(entry, "No")) || 0);
		});
		return Math.max(1, maxNo + 1, childElements(root, "Entry").length + 1);
	}

	function initState() {
		var mode = isEditMode() ? "edit" : "create";
		var voucherRoot = xmlRoot("VoucherData");
		window.VoucherRoot = voucherRoot;
		window.VouRoot = isEditMode() ? selectNodes(voucherRoot, "//Details")[0] || voucherRoot : xmlRoot("DetData");
		window.EntryRoot = window.EntryRoot && window.EntryRoot.nodeType ? window.EntryRoot : xmlRoot("EntryData");
		if (!window.iEntryNo || window._PurchaseVoucherEntryMode !== mode) {
			window.iEntryNo = nextEntryNo(window.VouRoot);
		}
		window._PurchaseVoucherEntryMode = mode;
		return { vouRoot: window.VouRoot, entryRoot: window.EntryRoot, voucherRoot: voucherRoot };
	}

	function resetEntryFields() {
		setValue("txtDescription", "");
		setValue("txtQty", "0.00");
		setValue("txtAmount", "0.00");
		setValue("txtDisAmount", "0.00");
		setValue("txtDisPercentage", "0.00");
		setValue("txtRate", "0.00");
		setValue("txtValue", "0.00");
		setValue("hItemCode", "0");
		setValue("hClassCode", "0");
		setValue("hEditEntNo", "0");
		if (trim(valueOf("hSalAccCode")) === "0" && field("selAccountHead")) {
			field("selAccountHead").selectedIndex = 0;
			setText("spAccHead", "");
		} else if (field("selAccountHead")) {
			selectByValue(field("selAccountHead"), "G");
			setText("spAccHead", valueOf("hSalAccName"));
		}
	}

	function appendAccHeadFromData() {
		var state = initState();
		childElements(xmlRoot("AccHeadData")).forEach(function (node) {
			state.entryRoot.appendChild(importFor(state.entryRoot, node));
		});
	}

	function updateEntryAnalysisAmounts(entryRoot) {
		childElements(entryRoot).forEach(function (header) {
			childElements(header).forEach(function (node) {
				var code = attr(node, "No") || attr(node, 0);
				var groupCode = attr(node, "GroupCode");
				var ratio;
				var amount;
				if (header.nodeName === "CostCenter") {
					ratio = field("txtCCRatio" + code);
					amount = field("txtCCAmount" + code);
				}
				if (header.nodeName === "Analytical") {
					ratio = field("txtANALRatio" + code + "Z" + groupCode);
					amount = field("txtANALAmount" + code + "Z" + groupCode);
				}
				if (ratio) {
					setAttr(node, "Ratio", ratio.value);
				}
				if (amount) {
					setAttr(node, "Amount", amount.value);
				}
			});
		});
	}

	function updateDetailsTotals(detailsRoot) {
		var total = 0;
		if (!detailsRoot) {
			return;
		}
		childElements(detailsRoot, "Entry").forEach(function (entry) {
			total += toNumber(attr(entry, "Amount"));
		});
		setAttr(detailsRoot, "BasicValue", formatNumber(total, 2));
		setAttr(detailsRoot, "ActualValue", formatNumber(total, 2));
		setAttr(detailsRoot, "VouDate", getDateControl("ctlDate"));
	}

	function removeChildrenByName(root, name) {
		childElements(root, name).forEach(function (node) {
			root.removeChild(node);
		});
	}

	function updateSingleNode(root, expression, updater) {
		var node = selectNodes(root, expression)[0];
		if (node) {
			updater(node);
		}
	}

	function syncVoucherHeaderFromForm() {
		var root = xmlRoot("VoucherData");
		var book = field("selBook");
		var purType = field("selPurType");
		var invoiceDate = getDateControl("ctlDate");
		var partyParts;
		if (!window.validate()) {
			return false;
		}
		setValue("hSetInvDate", invoiceDate);
		childElements(xmlRoot("UnitBookData")).forEach(function (bookNode) {
			if (book && attr(bookNode, 0) === book.value) {
				updateSingleNode(root, "//Book", function (node) {
					setAttr(node, "BookId", book.value);
					setAttr(node, "BKAccHead", attr(bookNode, 2));
					setAttr(node, "BKOtherUnits", attr(bookNode, 3));
					node.textContent = selectedText(book);
				});
			}
		});
		updateSingleNode(root, "//PurchaseType", function (node) {
			setAttr(node, "PurTypeId", purType ? purType.value : "");
			node.textContent = selectedText(purType);
		});
		updateSingleNode(root, "//PurInvoice", function (node) {
			setAttr(node, "PurInvNo", valueOf("txtInvoiceNo"));
			setAttr(node, "PurInvDate", invoiceDate);
		});
		partyParts = trim(valueOf("hPartyCode")).replace(/\|/g, "?").split("?");
		updateSingleNode(root, "//Party", function (node) {
			setAttr(node, "ParType", partyParts[0] || "");
			setAttr(node, "ParSubType", partyParts[1] || "");
			setAttr(node, "ParSubTypeName", selectedText(field("selPartyType")) || partyParts[2] || "");
			setAttr(node, "ParCode", partyParts[3] || "");
			node.textContent = valueOf("txtPartyName");
		});
		updateSingleNode(root, "//Details", updateDetailsTotals);
		return true;
	}

	function syncAmdDetailsHeaderFromForm() {
		var root = xmlRoot("VoucherData");
		var invoiceNo = valueOf("txtInvoiceNo") || textOf("tInvNo");
		var invoiceDate = valueOf("hSetInvDate");
		var voucherDate = getDateControl("ctlDate") || valueOf("hVouDate");
		setValue("hVouDate", voucherDate);
		updateSingleNode(root, "//PurInvoice", function (node) {
			setAttr(node, "PurInvNo", invoiceNo);
			setAttr(node, "PurInvDate", invoiceDate);
		});
		updateSingleNode(root, "//Details", function (node) {
			updateDetailsTotals(node);
			setAttr(node, "VouDate", voucherDate);
		});
		return true;
	}

	function setAmdDetailsDate() {
		var fromYear = trim(valueOf("hFromYr"));
		var toYear = trim(valueOf("hToYr"));
		var control = field("ctlDate") || byId("ctlDate");
		var minDate = fromYear ? "01/04/" + fromYear : "";
		var maxDate = toYear ? "31/03/" + toYear : "";
		if (control && minDate) {
			if (typeof control.SetMinDate === "function") {
				control.SetMinDate(minDate);
			} else if (typeof control.setMinDate === "function") {
				control.setMinDate(minDate);
			}
		}
		if (control && maxDate) {
			if (typeof control.SetMaxDate === "function") {
				control.SetMaxDate(maxDate);
			} else if (typeof control.setMaxDate === "function") {
				control.setMaxDate(maxDate);
			}
		}
		setDateControl("ctlDate", valueOf("hVouDate") || getDateControl("ctlDate") || valueOf("hSetInvDate"));
		setValue("hVouDate", getDateControl("ctlDate"));
	}

	function removeAmdTransientNodes() {
		var root = xmlRoot("VoucherData");
		removeChildrenByName(root, "TaxDetails");
		removeChildrenByName(root, "AdvanceDetails");
	}

	function hydrateHeaderFromVoucher() {
		var root = xmlRoot("VoucherData");
		var bookNode = selectNodes(root, "//Book")[0];
		var purTypeNode = selectNodes(root, "//PurchaseType")[0];
		var invoiceNode = selectNodes(root, "//PurInvoice")[0];
		var partyNode = selectNodes(root, "//Party")[0];
		var partyValue;
		if (bookNode) {
			selectByValue(field("selBook"), attr(bookNode, "BookId"));
			if (field("selBook") && field("selBook").selectedIndex > 0) {
				window.GetAccHead(field("selBook"));
			}
		}
		if (purTypeNode) {
			selectByValue(field("selPurType"), attr(purTypeNode, "PurTypeId"));
		}
		if (invoiceNode) {
			setValue("txtInvoiceNo", attr(invoiceNode, "PurInvNo") || valueOf("txtInvoiceNo"));
			setDateControl("ctlDate", attr(invoiceNode, "PurInvDate") || valueOf("hSetInvDate"));
			setValue("hSetInvDate", getDateControl("ctlDate"));
		}
		if (partyNode) {
			partyValue = attr(partyNode, "ParType") + "?" + attr(partyNode, "ParSubType");
			setValue("txtPartyName", partyNode.textContent || valueOf("txtPartyName"));
			setValue("hPartyCode", partyValue + "?" + (attr(partyNode, "ParSubTypeName") || "") + "?" + attr(partyNode, "ParCode"));
			selectByValue(field("selPartyType"), partyValue);
		}
	}

	window.clearXML = function () {
		var entry = createNode("EntryData", "Entry");
		setAttr(entry, "No", window.iEntryNo || 1);
		setAttr(entry, "PayTo", "");
		setAttr(entry, "Amount", "");
		setAttr(entry, "Qty", "");
		setAttr(entry, "UOM", "");
		setAttr(entry, "UOMValue", "");
		setAttr(entry, "Rate", "");
		setAttr(entry, "ActValue", "");
		setAttr(entry, "DisPer", "");
		setAttr(entry, "DisAmount", "");
		setAttr(entry, "ItemCode", "");
		setAttr(entry, "ClassCode", "");
		window.EntryRoot = entry;
		return entry;
	};

	window.CheckAccHead = function (root, accHead) {
		return selectNodes(root, "//AccHead[@No='" + accHead + "']").length > 0;
	};

	window.GetAccHead = function (select) {
		var xhr;
		var parts;
		var account = field("selAccountHead");
		if (!select || trim(select.value) === "S") {
			return false;
		}
		setValue("hBookcode", select.value);
		xhr = syncGet("GetAccHeadName.asp?Mod=PUR&BookNo=" + encodeURIComponent(select.value));
		parts = trim(xhr.responseText).split(":");
		if (parts.length > 1) {
			setValue("hSalAccCode", parts[0]);
			setValue("hSalAccName", parts[1]);
		}
		if (account) {
			selectByValue(account, trim(valueOf("hSalAccCode")) === "0" ? "S" : "G");
			account.disabled = true;
		}
		setText("spAccHead", valueOf("hSalAccName"));
		return false;
	};

	window.DisplayBook = function () {
		var select = field("selBook");
		var xhr;
		var desired;
		initState();
		if (select) {
			select.options.length = 1;
		}
		xhr = syncGet("XMLGetOrgBook.asp?BkCode=04&orgID=" + encodeURIComponent(valueOf("hOrgId")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("UnitBookData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("UnitBookData", xhr.responseText);
		}
		childElements(xmlRoot("UnitBookData")).forEach(function (node) {
			if (select) {
				select.add(new Option(attr(node, 1), attr(node, 0)));
			}
		});
		desired = valueOf("hBookNo") || valueOf("hBookcode") || attr(selectNodes(xmlRoot("VoucherData"), "//Book")[0], "BookId");
		if (desired) {
			selectByValue(select, desired);
		}
		if (select && select.selectedIndex > 0) {
			window.GetAccHead(select);
		}
		window.popPartyType();
	};

	window.popPartyType = function () {
		var select = field("selPartyType");
		var xhr;
		var hiddenParty = trim(valueOf("hPartyCode")).replace(/\|/g, "?").split("?");
		var hiddenType = valueOf("hPartyType");
		if (!select) {
			return;
		}
		select.options.length = 1;
		xhr = syncGet("../../Common/PartySubType.asp?ParCode=&OrgCode=" + encodeURIComponent(valueOf("hOrgId")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("OutData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("OutData", xhr.responseText);
		}
		childElements(xmlRoot("OutData")).forEach(function (node) {
			select.add(new Option(trim(node.textContent || node.text || ""), attr(node, "SubType")));
		});
		if (hiddenParty[0] && hiddenParty[1]) {
			selectByValue(select, hiddenParty[0] + "?" + hiddenParty[1]);
		}
		if (select.selectedIndex <= 0 && hiddenType) {
			for (var i = 0; i < select.options.length; i += 1) {
				if (trim(select.options[i].value).replace(/\|/g, "?").split("?")[0] === hiddenType) {
					select.selectedIndex = i;
					break;
				}
			}
		}
	};

	window.selAccHead = function (select) {
		var size;
		var partyType;
		var base;
		var url;
		select = select || field("selPartyType");
		if (!select || select.selectedIndex <= 0) {
			setValue("txtPartyName", "");
			setValue("hPartyCode", "0");
			return false;
		}
		size = popupSize("2", "PartySelection.asp", "500", "420");
		partyType = trim(select.value).replace(/\|/g, "?") + "?" + selectedText(select).replace(/&/g, " and ");
		base = "../../Common/" + size.program;
		url = base + "?orgid=" + encodeURIComponent(valueOf("hOrgId")) + "&Party=" + encodeURIComponent(partyType);
		runXmlDialog(url, base + "?", xmlObject("PartyData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (root) {
			var selected = childElements(root)[0];
			var partyName;
			var parCode;
			var accountNode;
			if (!selected) {
				select.selectedIndex = 0;
				setValue("txtPartyName", "");
				setValue("hPartyCode", "0");
				return;
			}
			partyName = attr(selected, "RetField0");
			parCode = attr(selected, "RetField1");
			clearChildren(xmlRoot("AccHeadData"));
			if (typeof window.GetPartyHeadXml === "function") {
				window.GetPartyHeadXml(parCode, partyName, "0:0:0");
			}
			accountNode = childElements(xmlRoot("AccHeadData"))[0];
			if (accountNode) {
				setValue("hPartyCode", partyType + "?" + (attr(accountNode, "No") || attr(accountNode, 0)));
				setValue("txtPartyName", attr(accountNode, "Name") || attr(accountNode, 3) || partyName);
			} else {
				select.selectedIndex = 0;
			}
		});
		return false;
	};

	window.showGLHead = function () {
		var size = popupSize("5", "GLHeadSelection.asp", "500", "350");
		var base = "../../Common/" + size.program;
		var url = base + "?orgID=" + encodeURIComponent(valueOf("hOrgId")) + "&BookId=01&BookNo=" + encodeURIComponent(valueOf("hBookcode")) + "&AccHead=" + encodeURIComponent(valueOf("hSalAccCode"));
		runXmlDialog(url, base + "?", xmlObject("GLHeadData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (root) {
			var ret = "";
			var accNode;
			childElements(root).forEach(function (node) {
				ret = [0, 1, 2, 3, 4, 5, 6, 7].map(function (index) {
					return attr(node, "RetField" + index);
				}).join(":");
			});
			clearChildren(xmlRoot("AccHeadData"));
			if (ret && typeof window.GetGlHeadXml === "function") {
				window.GetGlHeadXml(ret);
			}
			accNode = childElements(xmlRoot("AccHeadData"))[0];
			if (!accNode) {
				if (field("selAccountHead")) {
					field("selAccountHead").selectedIndex = 0;
				}
				setValue("txtDescription", "");
				setText("spAccHead", "");
				if (window.EntryRoot && window.EntryRoot.firstChild) {
					window.EntryRoot.removeChild(window.EntryRoot.firstChild);
				}
				if (typeof window.setADDDisplay === "function") {
					window.setADDDisplay(0);
				}
				return;
			}
			window.clearXML();
			window.EntryRoot.appendChild(importFor(window.EntryRoot, accNode));
			setText("spAccHead", attr(accNode, "Name") || attr(accNode, 3));
			setValue("txtDescription", "");
			if (typeof window.showCCAnal === "function") {
				window.showCCAnal(valueOf("hOrgId"), attr(accNode, "No") || attr(accNode, 0), attr(accNode, "CostCenter") || attr(accNode, 1), attr(accNode, "Analytical") || attr(accNode, 2));
			}
		});
		return false;
	};

	window.popSalesHead = function (select) {
		if (!select || select.selectedIndex <= 0) {
			return false;
		}
		if (select.value === "G") {
			return window.showGLHead();
		}
		setText("spAccHead", valueOf("hSalAccName"));
		return false;
	};

	window.checkFileds = function () {
		if (trim(valueOf("txtDescription")) === "" && trim(valueOf("hSalAccCode")) === "0") {
			alert("Select Item Description ");
			return false;
		}
		if (trim(valueOf("txtQty")) === "" || isNaN(Number(valueOf("txtQty")))) {
			alert(trim(valueOf("txtQty")) === "" ? "Enter Quantity" : "Enter Numeric values for Quantity");
			if (field("txtQty")) {
				field("txtQty").select();
			}
			return false;
		}
		if (trim(valueOf("txtRate")) === "" || isNaN(Number(valueOf("txtRate")))) {
			alert(trim(valueOf("txtRate")) === "" ? "Enter Rate" : "Enter Numeric values for Rate");
			if (field("txtRate")) {
				field("txtRate").select();
			}
			return false;
		}
		if (trim(valueOf("txtDisAmount")) === "" || isNaN(Number(valueOf("txtDisAmount")))) {
			alert(trim(valueOf("txtDisAmount")) === "" ? "Enter Discount" : "Enter Numeric values for Discount");
			if (field("txtDisAmount")) {
				field("txtDisAmount").select();
			}
			return false;
		}
		return true;
	};

	window.calculateField = function (flag) {
		var qty = toNumber(valueOf("txtQty"));
		var rate = toNumber(valueOf("txtRate"));
		var value;
		var discountAmount;
		if (!window.checkFileds()) {
			return false;
		}
		if (rate === 0) {
			return true;
		}
		value = rate * qty;
		setValue("txtValue", formatNumber(value, 2));
		if (Number(flag) === 2) {
			if (toNumber(valueOf("txtDisPercentage")) > 100) {
				alert("DisCount Percentage Should be less than 100");
				field("txtDisPercentage").select();
				return false;
			}
			discountAmount = value * (toNumber(valueOf("txtDisPercentage")) / 100);
			setValue("txtDisAmount", formatNumber(discountAmount, 2));
		} else if (Number(flag) === 3) {
			discountAmount = toNumber(valueOf("txtDisAmount"));
			if (discountAmount > value) {
				alert("DisCount Value Should be less than actual Value");
				field("txtDisAmount").select();
				return false;
			}
			setValue("txtDisPercentage", value ? formatNumber(discountAmount / value * 100, 2) : "0.00");
		} else {
			discountAmount = toNumber(valueOf("txtDisPercentage")) > 0 ? value * (toNumber(valueOf("txtDisPercentage")) / 100) : 0;
			setValue("txtDisAmount", formatNumber(discountAmount, 2));
		}
		setValue("txtAmount", formatNumber(value - toNumber(valueOf("txtDisAmount")), 2));
		window.popAddAmount1();
		return true;
	};

	window.popAddAmount1 = function () {
		var entryRoot = initState().entryRoot;
		var amount = toNumber(valueOf("txtAmount"));
		if (!window.checkFileds()) {
			setValue("txtAmount", "");
			return false;
		}
		childElements(entryRoot).forEach(function (header) {
			var nodes = childElements(header);
			var total = amount;
			var ratioTotal = 0;
			nodes.forEach(function (node, index) {
				var code = attr(node, "No") || attr(node, 0);
				var groupCode = attr(node, "GroupCode");
				var ratio;
				var splitAmount;
				var ratioField;
				var amountField;
				if (header.nodeName !== "CostCenter" && header.nodeName !== "Analytical") {
					return;
				}
				ratio = index < nodes.length - 1 ? Math.round((100 / nodes.length) * 100) / 100 : 100 - ratioTotal;
				splitAmount = index < nodes.length - 1 ? Math.round(((ratio * amount) / 100) * 100) / 100 : total;
				if (header.nodeName === "CostCenter") {
					ratioField = field("txtCCRatio" + code);
					amountField = field("txtCCAmount" + code);
				} else {
					ratioField = field("txtANALRatio" + code + "Z" + groupCode);
					amountField = field("txtANALAmount" + code + "Z" + groupCode);
				}
				if (ratioField) {
					ratioField.value = formatNumber(ratio, 2);
				}
				if (amountField) {
					amountField.value = formatNumber(splitAmount, 2);
				}
				setAttr(node, "Ratio", formatNumber(ratio, 2));
				setAttr(node, "Amount", formatNumber(splitAmount, 2));
				total -= splitAmount;
				ratioTotal += ratio;
			});
		});
		return true;
	};

	window.DisplayVoucher = function () {
		var root = initState().vouRoot;
		var table = clearTable("tblVoucher", 1, 1);
		var display = byId("DisVoucher");
		var total = 0;
		var row;
		if (display && display.style) {
			display.style.height = "200px";
			display.style.visibility = "visible";
		}
		if (!table) {
			return;
		}
		childElements(root, "Entry").forEach(function (entry, index) {
			var amount = isEditMode() ? toNumber(attr(entry, "ActValue")) - toNumber(attr(entry, "DisAmount")) : toNumber(attr(entry, "Amount"));
			if (isEditMode()) {
				setAttr(entry, "No", index + 1);
			}
			total += amount;
			row = table.insertRow(table.rows.length);
			insertCell(row, index + 1, "ExcelSerial", "Center", "top");
			insertCell(row, '<a href="javascript:EditEntry(\'' + escapeHtml(attr(entry, "No")) + '\')" class="ExcelDisplayCell"><b>Edit</b></a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, escapeHtml(attr(entry, "PayTo")), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(attr(entry, "Rate"), 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, escapeHtml(attr(entry, "Qty")) + "&nbsp;" + escapeHtml(attr(entry, "UOMValue")), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(attr(entry, "ActValue"), 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(attr(entry, "DisAmount"), 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(amount, 2), "ExcelDisplayCell", "right", "top");
		});
		row = table.insertRow(table.rows.length);
		insertCell(row, "<b>Total</b>", "ExcelDisplayCell", "right", "top", 7);
		insertCell(row, formatNumber(total, 2), "ExcelDisplayCell", "right", "top");
		window.iEntryNo = nextEntryNo(root);
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

	window.validate = function () {
		if (field("selBook") && field("selBook").selectedIndex < 1) {
			alert("Select Purchase Book");
			field("selBook").focus();
			return false;
		}
		if (field("selPurType") && field("selPurType").selectedIndex < 1) {
			alert("Select Purchase type");
			field("selPurType").focus();
			return false;
		}
		if (field("selPartyType") && field("selPartyType").selectedIndex < 1) {
			alert("Select Party");
			field("selPartyType").focus();
			return false;
		}
		if (trim(valueOf("txtPartyName")) === "") {
			alert("Party Name should not be blank");
			if (field("txtPartyName")) {
				field("txtPartyName").select();
			}
			return false;
		}
		if (trim(valueOf("txtInvoiceNo")) === "") {
			alert("Invoice No should not be blank");
			if (field("txtInvoiceNo")) {
				field("txtInvoiceNo").select();
			}
			return false;
		}
		return window.CheckVouStat();
	};

	window.VouCreate = function () {
		var root = xmlRoot("VoucherData");
		var header;
		var node;
		var book = field("selBook");
		var purType = field("selPurType");
		var invoiceDate = getDateControl("ctlDate");
		var duplicateValue;
		var xhr;
		var partyParts;
		if (!window.validate()) {
			return false;
		}
		if (!isEditMode()) {
			duplicateValue = valueOf("hPartyCode") + "?" + valueOf("txtInvoiceNo") + "?" + invoiceDate + "?04?" + valueOf("hOrgId");
			xhr = syncGet("CheckInvCreate.asp?sValue=" + encodeURIComponent(duplicateValue));
			if (trim(xhr.responseText) !== "C") {
				alert("Purchase Voucher already Created for this Party,InvoiceNo and Invoice Date ");
				return false;
			}
		}
		removeChildrenByName(root, "Header");
		header = createNode("VoucherData", "Header");
		root.appendChild(header);
		node = createNode("VoucherData", "Organization");
		setAttr(node, "OrgId", valueOf("hOrgId"));
		node.textContent = valueOf("hOrgName");
		header.appendChild(node);
		childElements(xmlRoot("UnitBookData")).forEach(function (bookNode) {
			var bookElement;
			if (!book || attr(bookNode, 0) !== book.value) {
				return;
			}
			bookElement = createNode("VoucherData", "Book");
			setAttr(bookElement, "BookId", book.value);
			setAttr(bookElement, "BKAccHead", attr(bookNode, 2));
			setAttr(bookElement, "BKOtherUnits", attr(bookNode, 3));
			bookElement.textContent = selectedText(book);
			header.appendChild(bookElement);
		});
		node = createNode("VoucherData", "PurchaseType");
		setAttr(node, "PurTypeId", purType ? purType.value : "");
		node.textContent = selectedText(purType);
		header.appendChild(node);
		node = createNode("VoucherData", "PurInvoice");
		setAttr(node, "PurInvNo", valueOf("txtInvoiceNo"));
		setAttr(node, "PurInvDate", invoiceDate);
		header.appendChild(node);
		node = createNode("VoucherData", "Party");
		partyParts = trim(valueOf("hPartyCode")).replace(/\|/g, "?").split("?");
		setAttr(node, "ParType", partyParts[0] || "");
		setAttr(node, "ParSubType", partyParts[1] || "");
		setAttr(node, "ParSubTypeName", selectedText(field("selPartyType")) || partyParts[2] || "");
		setAttr(node, "ParCode", partyParts[3] || "");
		node.textContent = valueOf("txtPartyName");
		header.appendChild(node);
		return true;
	};

	window.AddEntry = function (flag) {
		var state = initState();
		var entryRoot = state.entryRoot || window.clearXML();
		var amount = valueOf("txtAmount");
		var nextNo;
		if ((!isEditMode() || isAmdDetailsMode()) && trim(valueOf("hSalAccCode")) === "0" && field("selAccountHead") && field("selAccountHead").selectedIndex === 0) {
			alert("Select Purchase Account Head ");
			field("selAccountHead").focus();
			return false;
		}
		if (!childElements(entryRoot, "AccHead").length && toNumber(amount) !== 0) {
			clearChildren(xmlRoot("AccHeadData"));
			if (typeof window.GetGlHeadXmlForSalAcc === "function") {
				window.GetGlHeadXmlForSalAcc();
			}
			appendAccHeadFromData();
		}
		if (flag === "S" && childElements(state.vouRoot, "Entry").length > 0 && toNumber(amount) === 0) {
			window.SaveXML();
			return false;
		}
		if ((flag === "A" || flag === "U" || flag === "S") && !childElements(entryRoot).length) {
			alert("Select a Account Head");
			if (field("selAccountHead")) {
				field("selAccountHead").focus();
			}
			return false;
		}
		if (!window.checkFileds()) {
			return false;
		}
		setAttr(entryRoot, "No", flag === "U" ? valueOf("hEditEntNo") : window.iEntryNo);
		setAttr(entryRoot, "PayTo", valueOf("txtDescription"));
		setAttr(entryRoot, "Amount", formatNumber(valueOf("txtAmount"), 2));
		setAttr(entryRoot, "Qty", valueOf("txtQty"));
		setAttr(entryRoot, "UOM", valueOf("selUOM"));
		setAttr(entryRoot, "UOMValue", selectedText(field("selUOM")));
		setAttr(entryRoot, "Rate", formatNumber(valueOf("txtRate"), 2));
		setAttr(entryRoot, "ActValue", formatNumber(valueOf("txtValue"), 2));
		setAttr(entryRoot, "DisPer", formatNumber(valueOf("txtDisPercentage"), 2));
		setAttr(entryRoot, "DisAmount", formatNumber(valueOf("txtDisAmount"), 2));
		setAttr(entryRoot, "ItemCode", valueOf("hItemCode"));
		setAttr(entryRoot, "ClassCode", valueOf("hClassCode"));
		updateEntryAnalysisAmounts(entryRoot);
		if (flag === "U") {
			nextNo = Number(valueOf("hEditEntNo")) + 1;
			var nextNode = window._PurchaseVoucherEditNextNode || selectNodes(state.vouRoot, "//Entry[@No=" + nextNo + "]")[0];
			if (nextNode && nextNode.parentNode === state.vouRoot) {
				state.vouRoot.insertBefore(importFor(state.vouRoot, entryRoot), nextNode);
			} else {
				state.vouRoot.appendChild(importFor(state.vouRoot, entryRoot));
			}
			window._PurchaseVoucherEditNextNode = null;
		} else {
			state.vouRoot.appendChild(importFor(state.vouRoot, entryRoot));
		}
		updateDetailsTotals(state.vouRoot);
		if (flag === "S") {
			window.SaveXML();
			return false;
		}
		window.DisplayVoucher();
		window.iEntryNo = nextEntryNo(state.vouRoot);
		window.clearXML();
		resetEntryFields();
		setEntryButtons(false);
		if (typeof window.setADDDisplay === "function") {
			window.setADDDisplay(0);
		}
		return false;
	};

	window.EditEntry = function (entryNo) {
		var root = initState().vouRoot;
		var entry;
		if (trim(valueOf("hEditEntNo")) !== "0") {
			alert("Update the selected entry and edit this entry");
			return false;
		}
		entry = selectNodes(root, "//Entry[@No=" + entryNo + "]")[0];
		if (!entry) {
			return false;
		}
		setValue("txtDescription", attr(entry, "PayTo"));
		setValue("txtQty", attr(entry, "Qty"));
		setValue("txtRate", attr(entry, "Rate"));
		setValue("txtValue", attr(entry, "ActValue"));
		setValue("txtDisPercentage", attr(entry, "DisPer"));
		setValue("txtDisAmount", attr(entry, "DisAmount"));
		setValue("txtAmount", attr(entry, "Amount"));
		setValue("hItemCode", attr(entry, "ItemCode"));
		setValue("hClassCode", attr(entry, "ClassCode"));
		selectByValue(field("selUOM"), attr(entry, "UOM"));
		if (typeof window.setADDDisplay === "function") {
			window.setADDDisplay(0);
		}
		childElements(entry).forEach(function (node) {
			if (node.nodeName === "AccHead") {
				if (attr(node, "No") !== "0") {
					selectByValue(field("selAccountHead"), "G");
				}
				setText("spAccHead", attr(node, "Name"));
			}
			if (node.nodeName === "CostCenter" && typeof window.popCostCenter === "function") {
				window.setADDDisplay(1);
				window.popCostCenter(node);
			}
			if (node.nodeName === "Analytical" && typeof window.popAnalytical === "function") {
				window.setADDDisplay(1);
				window.popAnalytical(node);
			}
		});
		window._PurchaseVoucherEditNextNode = entry.nextSibling;
		window.EntryRoot = root.removeChild(entry);
		setValue("hEditEntNo", entryNo);
		setEntryButtons(true);
		window.DisplayVoucher();
		return false;
	};

	window.DelEntry = function () {
		window._PurchaseVoucherEditNextNode = null;
		window.clearXML();
		if (typeof window.setADDDisplay === "function") {
			window.setADDDisplay(0);
		}
		resetEntryFields();
		if (field("selUOM")) {
			field("selUOM").selectedIndex = 0;
		}
		if (field("selAccountHead")) {
			field("selAccountHead").selectedIndex = 0;
		}
		setEntryButtons(false);
		window.DisplayVoucher();
		return false;
	};

	window.GetItem = function () {
		var size = popupSize("1", "ItemSelectRelPartyCommon.asp", "500", "850");
		var base = "../../Common/" + size.program;
		runXmlDialog(base + "?orgID=" + encodeURIComponent(valueOf("hOrgId")), base + "?", xmlObject("ItemData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (root) {
			childElements(root).forEach(function (node) {
				setValue("txtDescription", attr(node, "ItemName"));
				setValue("hItemCode", attr(node, "ItemCode"));
				setValue("hClassCode", attr(node, "ClassCode"));
			});
		});
		return false;
	};

	window.SetDate = function () {
		var setDate = valueOf("hSetInvDate");
		var currentDate = new Date();
		var fromDate = valueOf("hFromDate");
		var toDate = valueOf("hToDate");
		var parsedTo = parseDate(toDate);
		var control = field("ctlDate") || byId("ctlDate");
		if (setDate) {
			setDateControl("ctlDate", setDate);
		} else if (parsedTo) {
			setDateControl("ctlDate", parsedTo.getTime() > currentDate.getTime() ? new Date() : toDate);
		}
		if (control) {
			if (fromDate && typeof control.SetMinDate === "function") {
				control.SetMinDate(fromDate);
			}
			if (typeof control.SetMaxDate === "function") {
				control.SetMaxDate(parsedTo && parsedTo.getTime() <= currentDate.getTime() ? toDate : new Date());
			}
		}
		setValue("hSetInvDate", getDateControl("ctlDate"));
	};

	window.SaveXML = function () {
		var xhr;
		if (!window.validate()) {
			return false;
		}
		if (isEditMode()) {
			if (isAmdDetailsMode()) {
				if (!window.CheckVouStat() || !syncAmdDetailsHeaderFromForm()) {
					return false;
				}
				xhr = syncPost("XMLSave.asp?Mod=PUR&Name=Voucher Amd", serializeXml("VoucherData"));
				if (trim(xhr.responseText) !== "") {
					alert(xhr.responseText);
					return false;
				}
				form().action = window.PurchaseVoucherAmdDetailsAction || "VouPURAmdTaxEntry.asp?sAmdType=A";
				form().submit();
				return true;
			}
			if (!syncVoucherHeaderFromForm()) {
				return false;
			}
			xhr = syncPost("XMLSave.asp?Mod=PUR&Name=Voucher Amd", serializeXml("VoucherData"));
			if (trim(xhr.responseText) !== "") {
				alert(xhr.responseText);
				return false;
			}
			form().submit();
			return true;
		}
		if (!window.VouCreate()) {
			return false;
		}
		xhr = syncPost("XMLSave.asp?Mod=PUR&Name=Voucher Entry", serializeXml("VoucherData"));
		xhr = syncPost("XMLUpdate.asp?Mod=PUR&Name=Voucher Entry", serializeXml("DetData"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		form().submit();
		return true;
	};

	window.SaveXML1 = function () {
		return window.SaveXML();
	};

	window.CancelAction = function (page) {
		form().action = page;
		form().submit();
	};

	window.InitPurchaseVoucherEntry = function () {
		window.PurchaseVoucherEntryMode = "create";
		window.PurchaseVoucherEntryVariant = "";
		ensureCompat();
		initState();
		window.SetDate();
		window.DisplayBook();
		setEntryButtons(false);
	};

	window.InitPurchaseVoucherEntryEdit = function () {
		window.PurchaseVoucherEntryMode = "edit";
		window.PurchaseVoucherEntryVariant = "";
		ensureCompat();
		initState();
		window.SetDate();
		window.DisplayBook();
		hydrateHeaderFromVoucher();
		window.DisplayVoucher();
		setEntryButtons(false);
	};

	window.InitPurchaseVoucherAmdDetails = function () {
		window.PurchaseVoucherEntryMode = "edit";
		window.PurchaseVoucherEntryVariant = "amd-details";
		ensureCompat();
		initState();
		setAmdDetailsDate();
		removeAmdTransientNodes();
		window.DisplayVoucher();
		setEntryButtons(false);
	};
}(window, document));
