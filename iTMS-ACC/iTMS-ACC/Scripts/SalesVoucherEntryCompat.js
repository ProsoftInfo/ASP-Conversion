(function (window, document) {
	"use strict";

	function isEditMode() {
		return window.SalesVoucherEntryMode === "edit";
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
		for (var i = 0; i < found.snapshotLength; i += 1) {
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

	function runStringDialog(url, nextPrefix, features, done) {
		openDialog(url, "", features, function (outValue) {
			var text = trim(outValue);
			var parts = text.split(":");
			if (!text) {
				return;
			}
			if (parts.length <= 1) {
				runStringDialog(nextPrefix + text, nextPrefix, features, done);
				return;
			}
			done(text, parts);
		});
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
		if (!window.iEntryNo || window._SalesVoucherEntryMode !== mode) {
			window.iEntryNo = nextEntryNo(window.VouRoot);
		}
		window._SalesVoucherEntryMode = mode;
		if (window.bVouFlag == null) {
			window.bVouFlag = false;
		}
		if (window.bSavFlag == null) {
			window.bSavFlag = false;
		}
		return { vouRoot: window.VouRoot, entryRoot: window.EntryRoot };
	}

	function populateSelect(select, root, textIndex, valueIndex) {
		if (!select) {
			return;
		}
		childElements(root).forEach(function (node) {
			select.add(new Option(attr(node, textIndex), attr(node, valueIndex)));
		});
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

	function selectByValue(select, value) {
		var wanted = trim(value).replace(/\|/g, "?");
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value).replace(/\|/g, "?") === wanted) {
				select.selectedIndex = i;
				return;
			}
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

	function setRadioValue(name, value) {
		fields(name).forEach(function (radio) {
			radio.checked = radio.value === value;
		});
	}

	function removeChildrenByName(root, name) {
		childElements(root, name).forEach(function (node) {
			root.removeChild(node);
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
		var saleType = field("selSaleType");
		var invoiceDate = getDateControl("ctlDate");
		var partyParts;
		if (!window.validate()) {
			return false;
		}
		childElements(xmlRoot("UnitBookData")).forEach(function (bookNode) {
			if (book && attr(bookNode, 0) === book.value) {
				setValue("hBkAccHead", attr(bookNode, 2));
				updateSingleNode(root, "//Book", function (node) {
					setAttr(node, "BookId", book.value);
					setAttr(node, "BKAccHead", attr(bookNode, 2));
					setAttr(node, "BKOtherUnits", attr(bookNode, 3));
					node.textContent = selectedText(book);
				});
			}
		});
		setValue("hInvDate", invoiceDate);
		updateSingleNode(root, "//Organization", function (node) {
			setAttr(node, "OrgId", valueOf("hOrgId"));
			node.textContent = valueOf("hOrgName");
		});
		updateSingleNode(root, "//SalesType", function (node) {
			setAttr(node, "SalType", saleType ? saleType.value : "");
			node.textContent = selectedText(saleType);
		});
		updateSingleNode(root, "//SaleInvoice", function (node) {
			setAttr(node, "InvNo", valueOf("txtInvoiceNo"));
			setAttr(node, "InvDate", invoiceDate);
			setAttr(node, "RefNo", valueOf("txtRefNo"));
		});
		updateSingleNode(root, "//Details", function (node) {
			setAttr(node, "VouDate", invoiceDate);
			updateDetailsTotals(node);
		});
		partyParts = trim(valueOf("hPartyCode")).replace(/\|/g, "?").split("?");
		updateSingleNode(root, "//Party", function (node) {
			setAttr(node, "ParType", partyParts[0] || "");
			setAttr(node, "ParSubType", partyParts[1] || "");
			setAttr(node, "ParSubTypeName", selectedText(field("selParType")) || partyParts[2] || "");
			setAttr(node, "ParCode", partyParts[3] || valueOf("hParCode"));
			setAttr(node, "Agent", selectedRadioValue("optAgentExist") === "C" ? "Y" : "N");
			node.textContent = valueOf("txtPartyName");
		});
		return true;
	}

	function populatePartySubTypes(parCode, selectedValue) {
		var select = field("selParType");
		var xhr;
		if (!select) {
			return;
		}
		select.options.length = 1;
		xhr = syncGet("../../Common/PartySubType.asp?ParCode=" + encodeURIComponent(parCode || "") + "&OrgCode=" + encodeURIComponent(valueOf("hOrgId")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("TEMPXML", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("TEMPXML", xhr.responseText);
		}
		childElements(xmlRoot("TEMPXML")).forEach(function (node) {
			select.add(new Option(trim(node.textContent || node.text || ""), attr(node, "SubType")));
		});
		selectByValue(select, selectedValue);
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

	function resetEntryFields() {
		setValue("hItemCode", "0");
		setValue("hClassCode", "0");
		setValue("txtDescription", "");
		setValue("txtQty", "0.00");
		setValue("txtBagno", "");
		setValue("txtAmount", "0.00");
		setValue("txtDisAmount", "0.00");
		setValue("txtDisPercentage", "0");
		setValue("txtRate", "0.00");
		setValue("txtRatePer", "1");
		setValue("txtValue", "0.00");
		setValue("hEditEntNo", "0");
		setValue("hEntryNo", "0");
		if (field("selUOM")) {
			setText("spUOM", selectedText(field("selUOM")));
		}
		if (field("selPack")) {
			setText("spPack", selectedText(field("selPack")));
		}
		if (trim(valueOf("hSalAccName")) === "" && field("selAccountHead")) {
			setText("spAccHead", "");
			field("selAccountHead").selectedIndex = 0;
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
		if (isEditMode()) {
			setAttr(entry, "ItemCode", "");
			setAttr(entry, "ClassCode", "");
			setAttr(entry, "TransBasicamt", "");
			setAttr(entry, "TransRate", "");
			setAttr(entry, "TransDisAmt", "");
			setAttr(entry, "TransInvAmt", "");
			setAttr(entry, "RatePer", "");
			setAttr(entry, "NoofPack", "");
			setAttr(entry, "PackType", "");
			setAttr(entry, "RndOff", "");
		} else {
			setAttr(entry, "RndOff", "");
			setAttr(entry, "NoofPack", "");
			setAttr(entry, "PackType", "");
			setAttr(entry, "RatePer", "");
			setAttr(entry, "ItemCode", "");
			setAttr(entry, "ClassCode", "");
		}
		window.EntryRoot = entry;
		return entry;
	};

	window.GetAccHead = function (select) {
		var xhr = syncGet("GetAccHeadName.asp?Mod=SAL&InvType=" + encodeURIComponent(select && select.value || ""));
		var parts = trim(xhr.responseText).split(":");
		var account = field("selAccountHead");
		if (parts.length > 1) {
			setValue("hSalAccCode", parts[0]);
			setValue("hSalAccName", parts[1]);
		}
		if (account) {
			account.disabled = true;
			selectByValue(account, trim(valueOf("hSalAccCode")) === "0" ? "S" : "G");
		}
		setText("spAccHead", trim(valueOf("hSalAccCode")) === "0" ? "" : valueOf("hSalAccName"));
	};

	window.popPartyType = function () {
		var select = field("selParType");
		var xhr;
		if (!select) {
			return;
		}
		select.options.length = 1;
		xhr = syncGet("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(valueOf("hOrgId")) + "&sCallTy=P");
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("OutData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("OutData", xhr.responseText);
		}
		childElements(xmlRoot("OutData")).forEach(function (node) {
			select.add(new Option(trim(node.textContent || node.text || ""), attr(node, "ParType") || attr(node, 0)));
		});
	};

	window.DisplayBook = function () {
		var select = field("selBook");
		var xhr;
		initState();
		if (select) {
			select.options.length = 1;
		}
		xhr = syncGet("XMLGetOrgBook.asp?BkCode=05&orgID=" + encodeURIComponent(valueOf("hOrgId")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("UnitBookData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("UnitBookData", xhr.responseText);
		}
		populateSelect(select, xmlRoot("UnitBookData"), 1, 0);
		window.popPartyType();
	};

	window.PopulateSalTy = function () {
		var book = field("selBook");
		var select = field("selSaleType");
		var xhr;
		if (!book || trim(book.value) === "S") {
			alert("Select Sales Book");
			return false;
		}
		setValue("hBookcode", book.value);
		if (select) {
			select.options.length = 1;
		}
		xhr = syncGet("XMLGetBookSalPurType.asp?BkCode=05&orgID=" + encodeURIComponent(valueOf("hOrgId")) + "&BookNo=" + encodeURIComponent(book.value));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("SaleTypeData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("SaleTypeData", xhr.responseText);
		}
		childElements(xmlRoot("SaleTypeData")).forEach(function (node) {
			if (select) {
				select.add(new Option(attr(node, 2), attr(node, 0)));
			}
		});
		return true;
	};

	function selectPartyNameForEdit() {
		var size = popupSize("2", "PartySelection.asp", "500", "420");
		var base = "../../Common/" + size.program;
		var url = base + "?orgid=" + encodeURIComponent(valueOf("hOrgId")) + "&Party=";
		runXmlDialog(url, base + "?", xmlObject("PartyData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (root) {
			var selected = childElements(root)[0];
			var parType;
			var parSubType;
			var parCode;
			var partyName;
			var select;
			var accountNode;
			if (!selected) {
				return;
			}
			parType = attr(selected, "RetField3");
			parSubType = attr(selected, "RetField4");
			parCode = attr(selected, "RetField1");
			partyName = attr(selected, "RetField0");
			setValue("hParCode", parCode);
			setValue("txtPartyName", partyName);
			populatePartySubTypes(parCode, parType + "|" + parSubType);
			select = field("selParType");
			if (select && select.selectedIndex <= 0) {
				selectByValue(select, parType + "?" + parSubType);
			}
			clearChildren(xmlRoot("AccHeadData"));
			if (typeof window.GetPartyHeadXml === "function") {
				window.GetPartyHeadXml(parCode, partyName, parType + ":" + parSubType + ":" + parCode + ":" + partyName + ":0");
			}
			accountNode = childElements(xmlRoot("AccHeadData"))[0];
			setValue("hPartyCode", parType + "?" + parSubType + "?" + selectedText(select) + "?" + (attr(accountNode, "No") || attr(accountNode, 0) || parCode));
		});
		return false;
	}

	window.SelPartyName = function () {
		var select = field("selParType");
		var size;
		var partyType;
		var url;
		if (isEditMode()) {
			return selectPartyNameForEdit();
		}
		if (!select || select.selectedIndex <= 0) {
			alert("Select Party Sub Type");
			if (select) {
				select.focus();
			}
			return false;
		}
		size = popupSize("12", "PartySelection.asp", "500", "500");
		partyType = select.value + "?" + selectedText(select);
		url = "../../Common/" + size.program + "?orgID=" + encodeURIComponent(valueOf("hOrgId")) + "&Party=" + encodeURIComponent(partyType);
		runStringDialog(url, "../../Common/" + size.program + "?", "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (outValue, parts) {
			var node;
			clearChildren(xmlRoot("AccHeadData"));
			setValue("hParCode", parts[1] || "");
			setValue("txtPartyName", parts[0] || "");
			if (typeof window.GetPartyHeadXml === "function") {
				window.GetPartyHeadXml(parts[1] || "", parts[0] || "", outValue + ":0");
			}
			node = childElements(xmlRoot("AccHeadData"))[0];
			if (node) {
				setValue("hPartyCode", partyType + "?" + attr(node, 0));
				setValue("txtPartyName", attr(node, 3) || parts[0] || "");
			} else {
				select.selectedIndex = 0;
				select.focus();
			}
		});
		return false;
	};

	window.PartyType = function (select) {
		var rawValue;
		var parts;
		var parCode;
		var partyName;
		var accountNode;
		if (!isEditMode()) {
			return window.SelPartyName();
		}
		select = select || field("selParType");
		rawValue = select ? select.value : "";
		if (!rawValue || rawValue === "A") {
			alert("Select Party Sub Type");
			if (select) {
				select.focus();
			}
			return false;
		}
		parts = trim(rawValue).replace(/\|/g, "?").split("?");
		parCode = valueOf("hParCode");
		partyName = valueOf("txtPartyName");
		clearChildren(xmlRoot("AccHeadData"));
		if (typeof window.GetPartyHeadXml === "function") {
			window.GetPartyHeadXml(parCode, partyName, (parts[0] || "") + ":" + (parts[1] || "") + ":" + parCode + ":" + partyName + ":0");
		}
		accountNode = childElements(xmlRoot("AccHeadData"))[0];
		if (accountNode) {
			setValue("hPartyCode", (parts[0] || "") + "?" + (parts[1] || "") + "?" + selectedText(select) + "?" + (attr(accountNode, "No") || attr(accountNode, 0)));
			setValue("txtPartyName", attr(accountNode, "Name") || attr(accountNode, 3) || partyName);
		} else if (select) {
			select.selectedIndex = 0;
		}
		return false;
	};

	window.PartyTypeForCreate = function () {
		return window.SelPartyName();
	};

	window.showGLHead = function () {
		var size = popupSize("5", "GLHeadSelection.asp", "500", "350");
		var base = "../../Common/" + size.program;
		var url = base + "?orgID=" + encodeURIComponent(valueOf("hOrgId")) + "&BookId=01&BookNo=" + encodeURIComponent(valueOf("hBookcode"));
		runXmlDialog(url, base + "?", xmlObject("GLHeadData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (root) {
			var ret = "";
			var accNode;
			childElements(root).forEach(function (node) {
				ret = [0, 1, 2, 3, 4, 5, 6, 7].map(function (index) {
					return attr(node, "RetField" + index);
				}).join(":");
			});
			if (!ret) {
				return;
			}
			clearChildren(xmlRoot("AccHeadData"));
			if (typeof window.GetGlHeadXml === "function") {
				window.GetGlHeadXml(ret);
			}
			accNode = childElements(xmlRoot("AccHeadData"))[0];
			if (!accNode) {
				setValue("txtDescription", "");
				if (typeof window.setADDDisplay === "function") {
					window.setADDDisplay(0);
				}
				return;
			}
			window.clearXML();
			window.EntryRoot.appendChild(importFor(window.EntryRoot, accNode));
			setText("spAccHead", attr(accNode, "Name") || attr(accNode, 3));
			if (typeof window.showCCAnal === "function") {
				window.showCCAnal(valueOf("hOrgId"), attr(accNode, "No") || attr(accNode, 0), attr(accNode, "CostCenter") || attr(accNode, 1), attr(accNode, "Analytical") || attr(accNode, 2));
			}
		});
		return false;
	};

	window.popSalesHead = function (objAcc) {
		var parts;
		var node;
		if (!objAcc || objAcc.selectedIndex <= 0) {
			return false;
		}
		if (objAcc.value === "G") {
			return window.showGLHead();
		}
		if (objAcc.value === "S") {
			return false;
		}
		parts = String(objAcc.value || "").split("?");
		node = createNode("EntryData", "AccHead");
		setAttr(node, "No", trim(parts[0]));
		setAttr(node, "CostCenter", trim(parts[1]));
		setAttr(node, "Analytical", trim(parts[2]));
		setAttr(node, "Name", selectedText(objAcc));
		setAttr(node, "Type", "G");
		setAttr(node, "Group", "");
		window.EntryRoot.appendChild(node);
		setText("spAccHead", selectedText(objAcc));
		if (typeof window.showCCAnal === "function") {
			window.showCCAnal(valueOf("hOrgId"), trim(parts[0]), trim(parts[1]), trim(parts[2]));
		}
		return false;
	};

	window.ValidateAmount = function (amount, name, from, to) {
		var number = toNumber(amount);
		if (trim(amount) === "") {
			alert(name + " Cannot be blank");
			return false;
		}
		if (isNaN(Number(String(amount).replace(/,/g, "")))) {
			alert("Enter Numeric values for " + name);
			return false;
		}
		if (number < Number(from) || number > Number(to)) {
			alert(name + " should be >" + from + " and < " + to);
			return false;
		}
		return true;
	};

	window.checkFileds = function () {
		if (!window.ValidateAmount(valueOf("txtQty"), "Quantity", 1, 9999999.999)) {
			field("txtQty").select();
			return false;
		}
		if (!window.ValidateAmount(valueOf("txtRate"), "Rate", 0, 9999999999.99)) {
			field("txtRate").select();
			return false;
		}
		if (!window.ValidateAmount(valueOf("txtDisAmount"), "Discount", 0, 9999999999.99)) {
			field("txtDisAmount").select();
			return false;
		}
		return true;
	};

	window.calculateField = function (flag) {
		var qty = toNumber(valueOf("txtQty"));
		var rate = toNumber(valueOf("txtRate"));
		var ratePer = toNumber(valueOf("txtRatePer"));
		var value;
		var discountPercent;
		var discountAmount;
		if (!window.ValidateAmount(valueOf("txtQty"), "Quantity", 0, 9999999.999)) {
			field("txtQty").select();
			return false;
		}
		if (!window.ValidateAmount(valueOf("txtRate"), "Rate", 0, 9999999999.99)) {
			field("txtRate").select();
			return false;
		}
		if (!window.ValidateAmount(valueOf("txtRatePer"), "Rate", 0.000001, 9999999999.99)) {
			field("txtRatePer").select();
			return false;
		}
		if (rate === 0) {
			return true;
		}
		value = rate / ratePer * qty;
		setValue("txtValue", formatNumber(value, 2));
		if (Number(flag) === 2) {
			discountPercent = toNumber(valueOf("txtDisPercentage"));
			if (discountPercent > 100) {
				alert("DisCount Percentage Should be less than 100");
				field("txtDisPercentage").select();
				return false;
			}
			discountAmount = value * (discountPercent / 100);
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
			discountPercent = toNumber(valueOf("txtDisPercentage"));
			discountAmount = discountPercent > 0 ? value * (discountPercent / 100) : toNumber(valueOf("txtDisAmount"));
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
				var ratioField;
				var amountField;
				var ratio;
				var splitAmount;
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
		var total = 0;
		var display = byId("DisVoucher");
		var row;
		if (display && display.style) {
			display.style.height = "200px";
			display.style.visibility = "visible";
		}
		if (!table) {
			return;
		}
		childElements(root, "Entry").forEach(function (entry, index) {
			var amount = toNumber(attr(entry, "Amount"));
			var description = attr(entry, "PayTo");
			var accountNode = childElements(entry, "AccHead")[0];
			if (isEditMode()) {
				setAttr(entry, "No", index + 1);
				if (accountNode && trim(attr(accountNode, "Name")) !== "") {
					description = attr(accountNode, "Name") + " - " + description;
				}
			}
			total += amount;
			row = table.insertRow(table.rows.length);
			insertCell(row, index + 1, "ExcelSerial", "Center", "top");
			insertCell(row, '<a href="#" onclick="EditEntry(\'' + escapeHtml(attr(entry, "No")) + '\'); return false;" class="ExcelDisplayCell"><b>Edit</b></a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, escapeHtml(description), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(attr(entry, "Qty"), 3) + "&nbsp;" + escapeHtml(attr(entry, "UOMValue")), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(attr(entry, "Rate"), 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(attr(entry, "ActValue"), 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(attr(entry, "DisAmount"), 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(amount, 2), "ExcelDisplayCell", "right", "top");
		});
		row = table.insertRow(table.rows.length);
		insertCell(row, "<b>Total</b>", "ExcelDisplayCell", "right", "top", 7);
		insertCell(row, formatNumber(total, 2), "ExcelDisplayCell", "right", "top");
	};

	window.validate = function () {
		var book = field("selBook");
		var saleType = field("selSaleType");
		var partyType = field("selParType");
		var partyName = field("txtPartyName");
		var invoiceNo = field("txtInvoiceNo");
		if (book && book.selectedIndex < 1) {
			alert("Select SalesBook");
			book.focus();
			return false;
		}
		if (saleType && saleType.selectedIndex < 1) {
			alert("Select Sales type");
			saleType.focus();
			return false;
		}
		if (partyType && partyType.selectedIndex < 1) {
			alert("Select Party");
			partyType.focus();
			return false;
		}
		if (partyName && trim(partyName.value) === "") {
			alert("Party Name should not be blank");
			partyName.select();
			return false;
		}
		if (invoiceNo && trim(invoiceNo.value) === "") {
			alert("Enter Invoice Number ");
			invoiceNo.focus();
			return false;
		}
		return true;
	};

	window.VouCreate = function () {
		var root = xmlRoot("VoucherData");
		var header;
		var node;
		var book = field("selBook");
		var saleType = field("selSaleType");
		var partyParts;
		var invoiceDate = getDateControl("ctlDate");
		var duplicateValue;
		var xhr;
		if (isEditMode()) {
			return syncVoucherHeaderFromForm();
		}
		if (!window.validate()) {
			return false;
		}
		duplicateValue = valueOf("hPartyCode") + "?" + valueOf("txtInvoiceNo") + "?" + invoiceDate + "?05?" + valueOf("hOrgId");
		xhr = syncGet("CheckInvCreate.asp?sValue=" + encodeURIComponent(duplicateValue));
		if (trim(xhr.responseText) !== "C") {
			alert("Sales Voucher already Created for this Party,InvoiceNo and Invoice Date ");
			return false;
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

		node = createNode("VoucherData", "SalesType");
		setAttr(node, "SalType", saleType ? saleType.value : "");
		node.textContent = selectedText(saleType);
		header.appendChild(node);

		node = createNode("VoucherData", "SaleInvoice");
		setAttr(node, "InvNo", valueOf("txtInvoiceNo"));
		setAttr(node, "InvDate", invoiceDate);
		setAttr(node, "RefNo", valueOf("txtRefNo"));
		header.appendChild(node);

		node = createNode("VoucherData", "Party");
		partyParts = trim(valueOf("hPartyCode")).split("?");
		setAttr(node, "ParType", partyParts[0] || "");
		setAttr(node, "ParSubType", partyParts[1] || "");
		setAttr(node, "ParSubTypeName", partyParts[2] || "");
		setAttr(node, "ParCode", partyParts[3] || "");
		setAttr(node, "Agent", selectedRadioValue("optAgentExist") === "C" ? "Y" : "N");
		node.textContent = valueOf("txtPartyName");
		header.appendChild(node);

		setAttr(xmlRoot("DetData"), "VouDate", invoiceDate);
		return true;
	};

	window.AddEntry = function (flag) {
		var state = initState();
		var entryRoot = state.entryRoot || window.clearXML();
		var appType = selectedRadioValue("optApproval", "Y");
		var amount = valueOf("txtAmount");
		var rounded = amount;
		var roundValue = 0;
		var nextNo;
		if (!window.validate()) {
			return false;
		}
		if (!childElements(entryRoot, "AccHead").length) {
			clearChildren(xmlRoot("AccHeadData"));
			if (typeof window.GetGlHeadXmlForSalAcc === "function") {
				window.GetGlHeadXmlForSalAcc();
			}
			appendAccHeadFromData();
		}
		if (flag === "S" && window.iEntryNo >= 1 && trim(amount) === "0.00") {
			if (isEditMode()) {
				updateDetailsTotals(state.vouRoot);
				window.SaveXML();
			} else if (window.VouCreate()) {
				window.SaveXML();
			}
			return false;
		}
		setAttr(entryRoot, "No", flag === "U" ? valueOf("hEditEntNo") : window.iEntryNo);
		if (!window.checkFileds()) {
			return false;
		}
		if (selectedRadioValue("optRound") === "Y") {
			rounded = typeof window.RndOff === "function" ? window.RndOff(amount) : Math.round(toNumber(amount));
			roundValue = Math.round((toNumber(rounded) - toNumber(amount)) * 100) / 100;
		}
		setAttr(entryRoot, "PayTo", valueOf("txtDescription"));
		setAttr(entryRoot, "Amount", rounded);
		setAttr(entryRoot, "Qty", valueOf("txtQty"));
		setAttr(entryRoot, "UOM", valueOf("selUOM"));
		setAttr(entryRoot, "UOMValue", selectedText(field("selUOM")));
		setAttr(entryRoot, "Rate", valueOf("txtRate"));
		setAttr(entryRoot, "ActValue", valueOf("txtValue"));
		setAttr(entryRoot, "DisPer", valueOf("txtDisPercentage"));
		setAttr(entryRoot, "DisAmount", valueOf("txtDisAmount"));
		setAttr(entryRoot, "RndOff", roundValue);
		setAttr(entryRoot, "NoofPack", valueOf("txtBagno"));
		setAttr(entryRoot, "PackType", valueOf("selPack"));
		setAttr(entryRoot, "RatePer", valueOf("txtRatePer"));
		setAttr(entryRoot, "ItemCode", valueOf("hItemCode"));
		setAttr(entryRoot, "ClassCode", valueOf("hClassCode"));
		if (isEditMode()) {
			setAttr(entryRoot, "TransBasicamt", valueOf("txtValue"));
			setAttr(entryRoot, "TransRate", valueOf("txtRate"));
			setAttr(entryRoot, "TransDisAmt", valueOf("txtDisAmount"));
			setAttr(entryRoot, "TransInvAmt", "0");
		}
		updateEntryAnalysisAmounts(entryRoot);
		if (flag === "U") {
			nextNo = Number(valueOf("hEditEntNo")) + 1;
			var nextNode = isEditMode() ? window._SalesVoucherEditNextNode : selectNodes(state.vouRoot, "//Entry[@No=" + nextNo + "]")[0];
			if (nextNode && nextNode.parentNode === state.vouRoot) {
				state.vouRoot.insertBefore(importFor(state.vouRoot, entryRoot), nextNode);
			} else {
				state.vouRoot.appendChild(importFor(state.vouRoot, entryRoot));
			}
			window._SalesVoucherEditNextNode = null;
		} else {
			state.vouRoot.appendChild(importFor(state.vouRoot, entryRoot));
		}
		window.bSavFlag = true;
		if (flag === "S") {
			if (isEditMode()) {
				updateDetailsTotals(state.vouRoot);
				window.SaveXML();
			} else if (window.VouCreate()) {
				window.SaveXML();
			}
			return false;
		}
		window.DisplayVoucher();
		window.iEntryNo = childElements(state.vouRoot, "Entry").length + 1;
		window.clearXML();
		resetEntryFields();
		setEntryButtons(false);
		fields("optApproval").forEach(function (radio) {
			radio.checked = radio.value === appType;
		});
		if (typeof window.setADDDisplay === "function") {
			window.setADDDisplay(0);
		}
		return false;
	};

	window.EditEntry = function (entryNo) {
		var root = initState().vouRoot;
		var entry = selectNodes(root, "//Entry[@No=" + entryNo + "]")[0];
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
		setValue("txtRatePer", attr(entry, "RatePer"));
		setValue("txtBagno", attr(entry, "NoofPack"));
		setRadioValue("optRound", toNumber(attr(entry, "RndOff")) !== 0 ? "Y" : "N");
		selectByValue(field("selUOM"), attr(entry, "UOM"));
		selectByValue(field("selPack"), attr(entry, "PackType"));
		setText("spUOM", selectedText(field("selUOM")));
		setText("spPack", selectedText(field("selPack")));
		setValue("hItemCode", attr(entry, "ItemCode"));
		setValue("hClassCode", attr(entry, "ClassCode"));
		if (typeof window.setADDDisplay === "function") {
			window.setADDDisplay(0);
		}
		childElements(entry).forEach(function (node) {
			var accountSelect;
			if (node.nodeName === "AccHead") {
				accountSelect = field("selAccountHead");
				selectByValue(accountSelect, attr(node, "No"));
				if (accountSelect && (!accountSelect.options[accountSelect.selectedIndex] || accountSelect.options[accountSelect.selectedIndex].value !== attr(node, "No"))) {
					selectByValue(accountSelect, attr(node, "Type") || "G");
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
		window._SalesVoucherEditNextNode = entry.nextSibling;
		window.EntryRoot = root.removeChild(entry);
		setValue("hEditEntNo", entryNo);
		setValue("hEntryNo", entryNo);
		setEntryButtons(true);
		window.DisplayVoucher();
		return false;
	};

	window.DelEntry = function () {
		window._SalesVoucherEditNextNode = null;
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

	window.ChDisp = function (select) {
		if (!select) {
			return;
		}
		if (String(select.name || "").toUpperCase() === "SELUOM") {
			setText("spUOM", selectedText(select));
		} else {
			setText("spPack", selectedText(select));
		}
	};

	window.showAgent = function (flag) {
		var root = xmlRoot("VoucherData");
		var noAgent = fields("optAgentExist")[2];
		removeChildrenByName(root, "AgentDetails");
		if (String(flag) === "N") {
			setText("spAgentName", "");
			setValue("hCommName", "");
			return false;
		}
		openDialog("AgentCommisionEntry.asp?OrgID=" + encodeURIComponent(valueOf("hOrgId")) + "&AgentType=" + encodeURIComponent(flag), xmlObject("OutData"), "dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", function (value) {
			var returnedRoot = rootFromDialog(value);
			var agent;
			if (returnedRoot && childElements(returnedRoot).length) {
				root.appendChild(importFor(root, returnedRoot));
				agent = selectNodes(root, "//Agent")[0];
				setText("spAgentName", agent ? attr(agent, 1) : "");
				setValue("hCommName", agent ? attr(agent, 1) : "");
			} else if (noAgent) {
				noAgent.checked = true;
			}
		});
		return false;
	};

	window.SaveXML = function () {
		var xhr;
		if (isEditMode()) {
			if (!syncVoucherHeaderFromForm()) {
				return false;
			}
			xhr = syncPost("XMLSave.asp?Mod=SAL&Name=Voucher AMD", serializeXml("VoucherData"));
			if (trim(xhr.responseText) !== "") {
				alert(xhr.responseText);
				return false;
			}
			form().submit();
			return true;
		}
		xhr = syncPost("XMLSave.asp?Mod=SAL&Name=Voucher Entry", serializeXml("VoucherData"));
		xhr = syncPost("XMLUpdate.asp?Mod=SAL&Name=Voucher Entry", serializeXml("DetData"));
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
		form().submit();
		return true;
	};

	window.DispOldVal = function () {
		var root = xmlRoot("VoucherData");
		var bookNode = selectNodes(root, "//Book")[0];
		var partyNode = selectNodes(root, "//Party")[0];
		var invoiceNode = selectNodes(root, "//SaleInvoice")[0];
		var saleTypeNode = selectNodes(root, "//SalesType")[0];
		var agentNode = selectNodes(root, "//AgentDetails/Agent")[0];
		var partyValue;
		var agentType = "";
		if (!root) {
			return;
		}
		if (bookNode) {
			setValue("hBkAccHead", attr(bookNode, "BKAccHead"));
			selectByValue(field("selBook"), attr(bookNode, "BookId"));
			window.PopulateSalTy();
		}
		if (partyNode) {
			setValue("txtPartyName", partyNode.textContent || "");
			setValue("hParCode", attr(partyNode, "ParCode"));
			partyValue = attr(partyNode, "ParType") + "?" + attr(partyNode, "ParSubType");
			selectByValue(field("selParType"), partyValue);
			if (field("selParType") && field("selParType").selectedIndex <= 0) {
				populatePartySubTypes(attr(partyNode, "ParCode"), partyValue);
			}
			setValue("hPartyCode", partyValue + "?" + (selectedText(field("selParType")) || attr(partyNode, "ParSubTypeName")) + "?" + attr(partyNode, "ParCode"));
		}
		if (agentNode && attr(partyNode, "Agent") === "Y") {
			agentType = attr(agentNode, "PartyType");
			setText("spAgentName", attr(agentNode, "Agentname") || attr(agentNode, "AgentName") || attr(agentNode, 1));
			setValue("hCommName", attr(agentNode, "Agentname") || attr(agentNode, "AgentName") || attr(agentNode, 1));
		} else {
			setText("spAgentName", "");
			setValue("hCommName", "");
		}
		setRadioValue("optAgentExist", agentType === "CR" ? "C" : agentType === "DR" ? "D" : "No");
		if (invoiceNode) {
			setValue("txtInvoiceNo", attr(invoiceNode, "InvNo"));
			setDateControl("ctlDate", attr(invoiceNode, "InvDate"));
			setValue("hInvDate", attr(invoiceNode, "InvDate"));
			setValue("txtRefNo", attr(invoiceNode, "RefNo"));
		}
		if (saleTypeNode) {
			selectByValue(field("selSaleType"), attr(saleTypeNode, "SalType"));
		}
		window.iEntryNo = nextEntryNo(initState().vouRoot);
	};

	window.CancelAction = function (page) {
		form().action = page;
		form().submit();
	};

	window.InitSalesVoucherEntry = function () {
		window.SalesVoucherEntryMode = "create";
		ensureCompat();
		initState();
		window.DisplayBook();
		setEntryButtons(false);
	};

	window.InitSalesVoucherEntryEdit = function () {
		window.SalesVoucherEntryMode = "edit";
		ensureCompat();
		initState();
		window.DisplayBook();
		window.DispOldVal();
		window.DisplayVoucher();
		setEntryButtons(false);
	};
}(window, document));
