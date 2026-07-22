(function (window, document) {
	"use strict";

	var stockIcon = "../../assets/images/iTMS%20Icons/EntryIcon.gif";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements && frm.elements[name] || document.getElementById(name) || window[name] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function checkNumbers(value, allowDecimal) {
		return (allowDecimal ? /^[0-9]+(\.[0-9]+)?$/ : /^[0-9]+$/).test(trim(value));
	}

	function selectedValue(nameOrSelect) {
		var select = typeof nameOrSelect === "string" ? field(nameOrSelect) : nameOrSelect;
		if (!select || !select.options || select.selectedIndex < 0) {
			return "";
		}
		return trim(select.options[select.selectedIndex].value);
	}

	function selectedText(nameOrSelect) {
		var select = typeof nameOrSelect === "string" ? field(nameOrSelect) : nameOrSelect;
		if (!select || !select.options || select.selectedIndex < 0) {
			return "";
		}
		return trim(select.options[select.selectedIndex].text);
	}

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var year;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			year = Number(match[3]);
			if (year < 100) {
				year += 2000;
			}
			return new Date(year, Number(match[2]) - 1, Number(match[1]));
		}
		return null;
	}

	function dateDiffDays(start, end) {
		var startDate = parseDate(start);
		var endDate = parseDate(end);
		if (!startDate || !endDate) {
			return 0;
		}
		return Math.floor((endDate.getTime() - startDate.getTime()) / 86400000);
	}

	function validDate(value) {
		var date = parseDate(value);
		if (!date) {
			return false;
		}
		return pad2(date.getDate()) + "/" + pad2(date.getMonth() + 1) + "/" + date.getFullYear() === trim(value).replace(/[.-]/g, "/");
	}

	function datePickerValue() {
		var control = field("ctlCDDate");
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return trim(control.GetDate());
		}
		if (typeof control.getDate === "function") {
			return trim(control.getDate());
		}
		return trim(control.value);
	}

	function setDatePickerValue(value) {
		var control = field("ctlCDDate");
		if (!control) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else {
			control.value = value;
		}
	}

	function xmlIsland(id) {
		ensureCompat();
		return window[id] || document[id] || document.getElementById(id);
	}

	function xmlDocument(value) {
		ensureCompat();
		if (!value) {
			return null;
		}
		if (value.XMLDocument) {
			return value.XMLDocument;
		}
		if (value._doc) {
			return value._doc;
		}
		if (value.nodeType === 9) {
			return value;
		}
		return value.ownerDocument || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || (value.nodeType === 1 ? value : null);
	}

	function serializeXml(value) {
		var doc = xmlDocument(value);
		var root = xmlRoot(value);
		if (doc && typeof doc.xml === "string") {
			return doc.xml;
		}
		if (value && typeof value.xml === "string") {
			return value.xml;
		}
		if (window.XMLSerializer) {
			return new XMLSerializer().serializeToString(doc || root);
		}
		return "";
	}

	function loadXml(id, text) {
		var target = xmlIsland(id);
		if (!target) {
			return null;
		}
		if (typeof target.loadXML === "function") {
			target.loadXML(text || "<Root/>");
			return xmlRoot(target);
		}
		return null;
	}

	function loadRootIntoIsland(id, root) {
		if (root) {
			loadXml(id, serializeXml(root));
		}
	}

	function createElement(doc, name) {
		return doc.createElement(name);
	}

	function importForDocument(doc, node) {
		if (!doc || !node) {
			return null;
		}
		if (node.ownerDocument === doc) {
			return node.cloneNode(true);
		}
		return doc.importNode ? doc.importNode(node, true) : node.cloneNode(true);
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || String(children[i].nodeName).toLowerCase() === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function attrAny(node, names) {
		var value = "";
		for (var i = 0; i < names.length; i += 1) {
			value = getAttr(node, names[i]);
			if (value !== "") {
				return value;
			}
		}
		return "";
	}

	function setAttr(node, name, value) {
		if (node) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function detailsIsland() {
		return xmlIsland("OutSelectData") || xmlIsland("OutData");
	}

	function detailsRoot() {
		return xmlRoot(detailsIsland());
	}

	function detailsDoc() {
		return xmlDocument(detailsIsland()) || detailsRoot() && detailsRoot().ownerDocument;
	}

	function isCreateMode() {
		return !!xmlIsland("OutSelectData");
	}

	function itemNodes() {
		return elementChildren(detailsRoot(), "ITEMDETAILS");
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function syncPost(url, xmlText) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(xmlText || "");
		return xhr;
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function findItemNode(itemCode, classCode, entryNo) {
		var nodes = itemNodes();
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "ITEMCODE") === trim(itemCode) &&
					getAttr(nodes[i], "CLASSCODE") === trim(classCode) &&
					(!entryNo || getAttr(nodes[i], "ENTRYNO") === trim(entryNo))) {
				return nodes[i];
			}
		}
		return null;
	}

	function removeHeaderNodes() {
		var root = detailsRoot();
		var headers = elementChildren(root, "HEADER");
		for (var i = headers.length - 1; i >= 0; i -= 1) {
			root.removeChild(headers[i]);
		}
	}

	function removeChildrenByName(node, names) {
		var wanted = {};
		var children = elementChildren(node);
		names.forEach(function (name) {
			wanted[String(name).toLowerCase()] = true;
		});
		for (var i = children.length - 1; i >= 0; i -= 1) {
			if (wanted[String(children[i].nodeName).toLowerCase()]) {
				node.removeChild(children[i]);
			}
		}
	}

	function selectOptionByValue(name, value) {
		var select = field(name);
		var wanted = trim(value).toLowerCase();
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value).toLowerCase() === wanted) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function setIssueToFromSelection() {
		var value = selectedValue("selIssueTo");
		var parts;
		if (!value || value.toLowerCase() === "select" || value === "0") {
			return;
		}
		parts = value.split(":");
		if (field("hIssueToType")) {
			field("hIssueToType").value = parts[0] || "";
		}
		if (field("hIssueToCode")) {
			field("hIssueToCode").value = parts.length > 1 ? parts[1] : "";
		}
	}

	function selectIssueToFromHidden() {
		var type = field("hIssueToType") ? field("hIssueToType").value : "";
		var code = field("hIssueToCode") ? field("hIssueToCode").value : "";
		var select = field("selIssueTo");
		var wanted = trim(type).toLowerCase() === "party" ? trim(type) : trim(type) + (code ? ":" + trim(code) : "");
		if (!select || !wanted) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value).toLowerCase() === wanted.toLowerCase()) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function normalizeAttributeList(value) {
		var text = trim(value);
		return text === "0" || text.toLowerCase() === "null" ? "" : text;
	}

	function itemDetailExists(root, itemCode, classCode, entryNo) {
		var nodes = elementChildren(root, "ITEMDETAILS");
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "ITEMCODE") === trim(itemCode) &&
					getAttr(nodes[i], "CLASSCODE") === trim(classCode) &&
					(!entryNo || getAttr(nodes[i], "ENTRYNO") === trim(entryNo))) {
				return true;
			}
		}
		return false;
	}

	function setItemTypeForItem(itemCode) {
		var xhr;
		var root;
		if (!field("hItemType") || trim(field("hItemType").value) || !itemCode) {
			return;
		}
		xhr = syncGet("../../Include/GetItemTypeForItem.asp?ItemCode=" + encodeURIComponent(itemCode));
		if (xhr.responseText) {
			root = loadXml("ItemTypeData", xhr.responseText);
			if (root) {
				field("hItemType").value = getAttr(root, "ItemType");
			}
		}
	}

	function appendItemDetail(source, options) {
		var root = detailsRoot();
		var doc = detailsDoc();
		var itemCode = attrAny(source, ["ItemCode", "ITEMCODE", "ITMCODE", "RetField1", "RetField0"]);
		var classCode = attrAny(source, ["ClassCode", "CLASSCODE", "CLACODE", "RetField2", "RetField1"]);
		var entryNo = attrAny(source, ["EntryNo", "ENTRYNO", "ICOUNTER", "ItemEntNo"]) || String(elementChildren(root, "ITEMDETAILS").length + 1);
		var itemName = attrAny(source, ["ItemName", "ITEMNAME", "ITMNAME", "RetField4"]);
		var uom = attrAny(source, ["StoresUoM", "StoresUOM", "UOM", "UoM", "RetField5"]);
		var decimal = attrAny(source, ["Decimal", "DECIMAL", "DecimalAllowed", "RetField6"]) || "0";
		var attrs = normalizeAttributeList(attrAny(source, ["AttributeList", "ATTRIBUTELIST", "ItemAttributes", "RetField8"]));
		var item;
		if (!root || !doc || !itemCode || !classCode || itemDetailExists(root, itemCode, classCode, entryNo)) {
			return;
		}
		item = createElement(doc, "ITEMDETAILS");
		setAttr(item, "ENTRYNO", entryNo);
		setAttr(item, "ITEMCODE", itemCode);
		setAttr(item, "CLASSCODE", classCode);
		setAttr(item, "UNIT", attrAny(source, ["UNIT", "OrgCode"]) || (field("hUnit") ? field("hUnit").value : ""));
		setAttr(item, "ITEMNAME", itemName.replace(/"/g, "~~"));
		setAttr(item, "UOM", uom);
		setAttr(item, "DECIMAL", decimal);
		setAttr(item, "DISPALYED", "N");
		setAttr(item, "DISPLAYED", "N");
		setAttr(item, "QTY", attrAny(source, ["Qty", "QTY", "QuantityRequested"]) || "0");
		setAttr(item, "REQUIREDBY", "");
		setAttr(item, "REQUIREDVALUE", "");
		setAttr(item, "ATTRIBUTELIST", attrs);
		setAttr(item, "REMARKS", "");
		setAttr(item, "RefNo", attrAny(source, ["No", "RefNo", "MRSNO"]));
		root.appendChild(item);
		if (options && options.detectItemType) {
			setItemTypeForItem(itemCode);
		}
	}

	function processReferenceSelection(selectionRoot) {
		var root = selectionRoot || xmlRoot(xmlIsland("OutData"));
		var refs = elementChildren(root, "Reference");
		var refCodes = [];
		var refNos = [];
		var refDates = [];
		var xhr;
		var itemRoot;
		if (!root) {
			return;
		}
		loadRootIntoIsland("OutData", root);
		refs.forEach(function (ref) {
			refCodes.push(getAttr(ref, "ReferenceCode") + " - " + getAttr(ref, "ReferenceDate"));
			refNos.push(getAttr(ref, "ReferenceNo"));
			refDates.push(getAttr(ref, "ReferenceDate"));
		});
		if (refNos.length) {
			if (field("RefNoDate")) {
				field("RefNoDate").innerHTML = refCodes.join(",");
			}
			if (field("hRefNo")) {
				field("hRefNo").value = refNos.join(",");
			}
			if (field("hRefDate")) {
				field("hRefDate").value = refDates.join(",");
			}
			xhr = syncGet("InvGetItemDetForRefType.asp?RefType=" + encodeURIComponent(selectedValue("selRefName")) + "&RefCodes=" + encodeURIComponent(refNos.join(",")) + "&orgID=" + encodeURIComponent(field("hUnit") ? field("hUnit").value : ""));
			if (xhr.responseText) {
				itemRoot = loadXml("ItemData", xhr.responseText);
				if (field("hItemType") && itemRoot && getAttr(itemRoot, "ItemType")) {
					field("hItemType").value = getAttr(itemRoot, "ItemType");
				}
				elementChildren(itemRoot).forEach(function (node) {
					if (/^item$/i.test(node.nodeName)) {
						appendItemDetail(node, { detectItemType: true });
					} else if (/^subcontract$/i.test(node.nodeName)) {
						detailsRoot().appendChild(importForDocument(detailsDoc(), node));
					}
				});
			}
		}
	}

	function processItemSelection(selectionRoot) {
		var root = selectionRoot || xmlRoot(xmlIsland("OutData"));
		if (!root) {
			return;
		}
		loadRootIntoIsland("OutData", root);
		elementChildren(root).forEach(function (node) {
			if (/^(item|row|data|returndata)$/i.test(node.nodeName) || attrAny(node, ["ItemCode", "ITEMCODE", "RetField1"])) {
				appendItemDetail(node, { detectItemType: true });
			}
		});
	}

	function appendCell(row, className, text) {
		var cell = row.insertCell(-1);
		cell.className = className || "ExcelDisplayCell";
		if (text != null) {
			cell.innerHTML = text;
		}
		return cell;
	}

	function appendInput(cell, name, value, size, readOnly) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.value = value == null ? "" : value;
		input.size = size || 12;
		input.className = readOnly ? "FormElemRead" : "FormElem";
		input.style.textAlign = readOnly ? "center" : "right";
		input.readOnly = !!readOnly;
		cell.appendChild(input);
		return input;
	}

	function addScheduleOptions(select, selected) {
		[
			["select", "Select"],
			["I", "Immediate"],
			["W", "Within x Days"],
			["D", "Specific Date"],
			["S", "Scheduled"]
		].forEach(function (entry) {
			var option = new Option(entry[1], entry[0]);
			if (entry[0] === selected) {
				option.selected = true;
			}
			select.options[select.options.length] = option;
		});
	}

	function renderItemRow(table, node, index) {
		var row = table.insertRow(-1);
		var entryNo = getAttr(node, "ENTRYNO");
		var itemCode = getAttr(node, "ITEMCODE");
		var classCode = getAttr(node, "CLASSCODE");
		var itemName = getAttr(node, "ITEMNAME").replace(/~~/g, "\"");
		var uom = getAttr(node, "UOM");
		var attrList = getAttr(node, "ATTRIBUTELIST");
		var cell;
		var select;
		var img;

		cell = appendCell(row, "ExcelDisplayCell");
		cell.align = "center";
		appendInput(cell, "txtSerA" + itemCode + "A" + classCode + "A" + entryNo, index, 1, true);

		cell = appendCell(row, "ExcelInputCell");
		cell.align = "center";
		cell.width = "10";
		img = document.createElement("input");
		img.type = "checkbox";
		img.name = "chkDeleteA" + itemCode + "A" + classCode + "A" + entryNo;
		img.value = entryNo;
		img.className = "FormElem";
		cell.appendChild(img);

		cell = appendCell(row, "ExcelDisplayCell", itemName);
		cell.align = "left";

		cell = appendCell(row, "ExcelInputCell");
		cell.width = "10";
		appendInput(cell, "txtQtyZ" + itemCode + "Z" + classCode + "Z" + entryNo, getAttr(node, "QTY") === "0" ? "" : getAttr(node, "QTY"), 12, false);

		cell = appendCell(row, "ExcelFieldCell");
		cell.align = "center";
		select = document.createElement("select");
		select.name = "selUoMZ" + itemCode + "Z" + classCode + "Z" + entryNo;
		select.className = "FormElem";
		select.options[0] = new Option(uom || "Select", uom || "select");
		cell.appendChild(select);

		cell = appendCell(row, "ExcelFieldCell");
		cell.align = "center";
		select = document.createElement("select");
		select.name = "selSchZ" + itemCode + "Z" + classCode + "Z" + entryNo;
		select.className = "FormElem";
		addScheduleOptions(select, getAttr(node, "REQUIREDBY"));
		select.onchange = function () {
			return window.CheckSch(select, datePickerValue(), uom, attrList);
		};
		cell.appendChild(select);

		cell = appendCell(row, "ExcelFieldCell");
		cell.align = "center";
		img = document.createElement("img");
		img.border = "0";
		img.src = stockIcon;
		img.alt = "Stock Details";
		img.width = 11;
		img.height = 11;
		img.style.cursor = "pointer";
		img.onclick = function () {
			return window.DisplayStock(itemCode, classCode, getAttr(node, "UNIT") || (field("hUnit") ? field("hUnit").value : ""), entryNo, itemName, attrList);
		};
		cell.appendChild(img);

		cell = appendCell(row, "ExcelInputCell");
		cell.width = "10";
		appendInput(cell, "txtRemZ" + itemCode + "Z" + classCode + "Z" + entryNo, getAttr(node, "REMARKS"), 12, false).style.textAlign = "left";
	}

	function setNodeDisplay(node, value) {
		setAttr(node, "DISPALYED", value);
		setAttr(node, "DISPLAYED", value);
	}

	function selectedIssueType() {
		return selectedValue("cmbIssType");
	}

	function currentApprover() {
		return selectedValue("selApprover");
	}

	function ensureScheduleNode(item, type, value, itemCode, classCode, entryNo) {
		var doc = detailsDoc();
		var schedule = createElement(doc, "Schedule");
		setAttr(schedule, "STYPE", type);
		setAttr(schedule, "SVALUE", value);
		setAttr(schedule, "ITEMCODE", itemCode);
		setAttr(schedule, "CLASSCODE", classCode);
		setAttr(schedule, "SCHENTRYNO", entryNo);
		item.appendChild(schedule);
		return schedule;
	}

	function getAttributeDisplayName(attrList) {
		var xhr;
		if (!trim(attrList)) {
			return "";
		}
		xhr = syncGet("XMLGetAttributeName.asp?Para=" + encodeURIComponent(attrList));
		return trim(xhr.responseText);
	}

	function validateEntryItems() {
		var nodes = itemNodes();
		var qty;
		var req;
		var uom;
		var remarks;
		var itemCode;
		var classCode;
		var entryNo;
		var i;
		if (!nodes.length) {
			alert("Select the Item(s)");
			return false;
		}
		for (i = 0; i < nodes.length; i += 1) {
			itemCode = getAttr(nodes[i], "ITEMCODE");
			classCode = getAttr(nodes[i], "CLASSCODE");
			entryNo = getAttr(nodes[i], "ENTRYNO");
			qty = field("txtQtyZ" + itemCode + "Z" + classCode + "Z" + entryNo);
			req = field("selSchZ" + itemCode + "Z" + classCode + "Z" + entryNo);
			uom = field("selUoMZ" + itemCode + "Z" + classCode + "Z" + entryNo);
			remarks = field("txtRemZ" + itemCode + "Z" + classCode + "Z" + entryNo);
			if (!qty || trim(qty.value) === "" || trim(qty.value) === "0") {
				alert("Enter Quantity");
				if (qty) {
					qty.select();
				}
				return false;
			}
			if (uom && trim(uom.value).toLowerCase() === "select") {
				alert("Select Unit Or Measurement");
				uom.focus();
				return false;
			}
			if (!req || req.selectedIndex === 0 || selectedValue(req).toLowerCase() === "select") {
				alert("Select Required By");
				if (req) {
					req.focus();
				}
				return false;
			}
			setAttr(nodes[i], "QTY", qty.value);
			setAttr(nodes[i], "UOM", uom ? uom.value : getAttr(nodes[i], "UOM"));
			setAttr(nodes[i], "REQUIREDBY", selectedValue(req));
			setAttr(nodes[i], "REMARKS", remarks ? trim(remarks.value) : getAttr(nodes[i], "REMARKS"));
		}
		return true;
	}

	function addCreateHeader() {
		var root = detailsRoot();
		var doc = detailsDoc();
		var header = createElement(doc, "HEADER");
		var refType = selectedValue("selRefName");
		setAttr(header, "FORUNIT", field("hUnit") ? field("hUnit").value : "");
		setAttr(header, "CREATEDON", datePickerValue());
		setAttr(header, "REMARKS", field("txtRemarks") ? trim(field("txtRemarks").value) : "");
		setAttr(header, "APPROVER", currentApprover());
		setAttr(header, "CREATEDBY", field("hCreatedBy") ? field("hCreatedBy").value : "");
		setAttr(header, "RECEIPTNO", "");
		setAttr(header, "COSTCENTER", selectedValue("selCC"));
		setAttr(header, "ITEMTYPE", field("hItemType") ? trim(field("hItemType").value) : "");
		setAttr(header, "ACCHEAD", selectedValue("selAccHead"));
		setAttr(header, "REFTYPE", "");
		setAttr(header, "AppRefType", refType && refType !== "N" ? refType : "");
		setAttr(header, "AppRefNo", field("hRefNo") ? field("hRefNo").value : "");
		setAttr(header, "AppRefDate", field("hRefDate") ? field("hRefDate").value : "");
		setAttr(header, "CallFrom", "MR");
		setAttr(header, "RedirectTo", "MRSMGMTLIST.ASP?HCHECK=M");
		setAttr(header, "ImmediateApprover", currentApprover() === "IM" ? "Y" : "N");
		setAttr(header, "MRNo", "");
		setAttr(header, "RequestedByUnit", field("hRequestedByUnit") ? field("hRequestedByUnit").value : "");
		setAttr(header, "ISSTOTYPE", field("hIssueToType") ? field("hIssueToType").value : "");
		setAttr(header, "ISSTOCODE", field("hIssueToCode") ? field("hIssueToCode").value : "");
		setAttr(header, "ISSTOSUBCODE", field("hIssueToSubCode") ? field("hIssueToSubCode").value : "");
		setAttr(header, "ISSUETYPECODE", selectedIssueType());
		root.appendChild(header);
	}

	function addApprovalHeader() {
		var root = detailsRoot();
		var doc = detailsDoc();
		var header = createElement(doc, "HEADER");
		setAttr(header, "FORUNIT", field("hUnit") ? field("hUnit").value : "");
		setAttr(header, "CREATEDON", datePickerValue());
		setAttr(header, "TYPE", "");
		setAttr(header, "USAGE", "");
		setAttr(header, "REMARKS", field("txtRemarks") ? trim(field("txtRemarks").value) : "");
		setAttr(header, "CREATEDBY", field("hCreatedBy") ? field("hCreatedBy").value : "");
		setAttr(header, "MRNO", field("hMRNo") ? field("hMRNo").value : "");
		setAttr(header, "LOTCARDNO", "");
		setAttr(header, "MACHINENO", "");
		setAttr(header, "COSTCENTER", selectedValue("selCC"));
		setAttr(header, "REFTYPE", "");
		setAttr(header, "ISSTOTYPE", field("hIssueToType") ? field("hIssueToType").value : "");
		setAttr(header, "ISSTOCODE", field("hIssueToCode") ? field("hIssueToCode").value : "");
		setAttr(header, "ISSTOSUBCODE", field("hIssueToSubCode") ? field("hIssueToSubCode").value : "");
		setAttr(header, "ISSUETYPECODE", selectedIssueType());
		root.appendChild(header);
	}

	window.Init = function (initialDate) {
		var fromDate;
		var toDate;
		var preferredDate;
		ensureCompat();
		fromDate = field("hFrmDate") ? field("hFrmDate").value : "";
		toDate = field("hToDate") ? field("hToDate").value : "";
		preferredDate = trim(initialDate) || (field("hMRDate") ? field("hMRDate").value : "") || toDate;
		if (preferredDate) {
			setDatePickerValue(preferredDate);
		} else if (toDate) {
			setDatePickerValue(toDate);
		}
		if (fromDate && toDate && (dateDiffDays(fromDate, datePickerValue()) < 0 || dateDiffDays(datePickerValue(), toDate) < 0)) {
			setDatePickerValue(toDate);
		}
		selectIssueToFromHidden();
		return false;
	};

	window.init = window.Init;

	window.MinDate = function () {
		var min = field("hFrmDate") ? field("hFrmDate").value : "";
		var max = field("hToDate") ? field("hToDate").value : "";
		var fallback = field("hMRDate") ? field("hMRDate").value : max;
		var value = datePickerValue();
		if (min && max && (dateDiffDays(min, value) < 0 || dateDiffDays(value, max) < 0)) {
			alert("Date Should be With in the Range " + min + " to " + max);
			setDatePickerValue(fallback);
		}
		return false;
	};

	window.popCC = function () {
		var select = field("selCC");
		var xhr;
		var root;
		if (!select) {
			return false;
		}
		select.length = 1;
		xhr = syncGet("XMLSelectCostCenter.asp?sOrgID=" + encodeURIComponent(field("hUnit") ? field("hUnit").value : ""));
		if (xhr.responseText) {
			root = loadXml("OutCost", xhr.responseText);
			elementChildren(root).forEach(function (node) {
				select.options[select.options.length] = new Option(attrAny(node, ["CName", "Description", "CostCenter", "Name"]) || (node.attributes && node.attributes[1] ? node.attributes[1].value : ""), attrAny(node, ["CCode", "Code", "CostCenterHead", "Value"]) || (node.attributes && node.attributes[0] ? node.attributes[0].value : ""));
			});
		}
		return false;
	};

	window.ClearTable = function () {
		var table = document.getElementById("tblLot");
		while (table && table.rows.length > 1) {
			table.deleteRow(1);
		}
		return false;
	};

	window.DisplayTable = function () {
		var table = document.getElementById("tblLot");
		var nodes = itemNodes();
		if (!table || !isCreateMode()) {
			return false;
		}
		window.ClearTable();
		nodes.forEach(function (node, index) {
			if (node.nodeName.toLowerCase() === "itemdetails") {
				renderItemRow(table, node, index + 1);
				setNodeDisplay(node, "Y");
			}
		});
		return false;
	};

	window.DeleteItems = function () {
		var root = detailsRoot();
		var nodes = itemNodes();
		var checkbox;
		var itemCode;
		var classCode;
		var entryNo;
		if (!root) {
			return false;
		}
		for (var i = nodes.length - 1; i >= 0; i -= 1) {
			itemCode = getAttr(nodes[i], "ITEMCODE");
			classCode = getAttr(nodes[i], "CLASSCODE");
			entryNo = getAttr(nodes[i], "ENTRYNO");
			checkbox = field("chkDeleteA" + itemCode + "A" + classCode + "A" + entryNo);
			if (isCreateMode()) {
				if (checkbox && checkbox.checked) {
					root.removeChild(nodes[i]);
				}
			} else {
				setNodeDisplay(nodes[i], checkbox && checkbox.checked ? "Y" : "N");
			}
		}
		if (isCreateMode()) {
			window.DisplayTable();
		}
		return false;
	};

	window.GetDetails = function () {
		var refType = selectedValue("selRefName");
		var orgId = field("hUnit") ? field("hUnit").value : "";
		if (typeof window.RefTypeSelection !== "function") {
			alert("Reference selection script is not available.");
			return false;
		}
		window.sUnit = orgId;
		window.sIType = field("hItemType") ? field("hItemType").value : "";
		window.RefTypeSelection(refType, orgId, "", "Y", 1, "N", 0, "MR", function (result) {
			if (trim(refType) !== "N") {
				processReferenceSelection(xmlRoot(result) || xmlRoot(xmlIsland("OutData")));
			} else {
				processItemSelection(xmlRoot(result) || xmlRoot(xmlIsland("OutData")));
			}
			window.DisplayTable();
		});
		return false;
	};

	window.GetItems = function () {
		var orgId = field("hUnit") ? field("hUnit").value : "";
		if (typeof window.RefTypeSelection === "function") {
			window.sUnit = orgId;
			window.sIType = field("hItemType") ? field("hItemType").value : "";
			window.RefTypeSelection("N", orgId, "", "Y", 1, "N", 0, "MR", function (result) {
				processItemSelection(xmlRoot(result) || xmlRoot(xmlIsland("OutData")));
				window.DisplayTable();
			});
		}
		return false;
	};

	window.GetAddDetails = function (itemCode, classCode, orgId, entryNo, attrList) {
		var qty = field("txtQtyZ" + itemCode + "Z" + classCode + "Z" + entryNo);
		var usage = field("selUsage") ? selectedValue("selUsage") : field("hIssueToCode") ? field("hIssueToCode").value : "";
		var tempValues;
		if (!qty || trim(qty.value) === "") {
			alert("Enter Quantity");
			return false;
		}
		if (toNumber(qty.value) === 0) {
			return false;
		}
		tempValues = classCode + "|" + itemCode + "|" + orgId + "|" + usage + "|" + trim(qty.value) + "|" + entryNo + "|" + (attrList || "");
		if (usage === "PRD") {
			openModal("DirectIssueMixEntry.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:300px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No");
		} else if (usage === "PAC") {
			openModal("DirectIssuePackingEntry.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:400px;dialogWidth:325px;center:Yes;help:No;resizable:No;status:No");
		} else {
			openModal("AddEntryDetails.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:370px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No");
		}
		return false;
	};

	window.DisplayStock = function (itemCode, classCode, orgId, entryNo, itemName) {
		openModal("../master/itmStockDetailsPop.asp?EntNo=" + encodeURIComponent(entryNo || "") + "&sItem=" + encodeURIComponent(itemCode || "") + "&sClass=" + encodeURIComponent(classCode || "") + "&ItemName=" + encodeURIComponent(itemName || "") + "&sOrg=" + encodeURIComponent(orgId || ""), "Stock", "dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No");
		return false;
	};

	window.CheckSch = function (obj, todaysDate, uom, attrList) {
		var parts = String(obj && obj.name || "").split("Z");
		var itemCode = parts[1] || "";
		var classCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var item = findItemNode(itemCode, classCode, entryNo);
		var selected = selectedValue(obj);
		var value;
		var qty;
		var optName;
		var tempValues;
		if (!item) {
			return false;
		}
		removeChildrenByName(item, ["Schedule", "ScheduleDetails"]);
		if (selected === "I") {
			ensureScheduleNode(item, "I", isCreateMode() ? datePickerValue() : todaysDate, itemCode, classCode, entryNo);
		} else if (selected === "W") {
			value = prompt("Enter No of Days", "0");
			if (value == null || trim(value) === "") {
				obj.selectedIndex = 0;
				return false;
			}
			if (!checkNumbers(value, false)) {
				alert("Enter Numerals Only");
				obj.selectedIndex = 0;
				return false;
			}
			ensureScheduleNode(item, "W", trim(value), itemCode, classCode, entryNo);
		} else if (selected === "D") {
			value = prompt("Enter the Date", "");
			if (value == null || trim(value) === "") {
				obj.selectedIndex = 0;
				return false;
			}
			if (!validDate(value) || dateDiffDays(todaysDate, value) < 0) {
				alert("Invalid Date");
				obj.selectedIndex = 0;
				return false;
			}
			ensureScheduleNode(item, "D", trim(value), itemCode, classCode, entryNo);
		} else if (selected === "S") {
			qty = field("txtQtyZ" + itemCode + "Z" + classCode + "Z" + entryNo);
			if (!qty || trim(qty.value) === "") {
				alert("Enter Quantity");
				if (qty) {
					qty.focus();
				}
				obj.selectedIndex = 0;
				return false;
			}
			if (!checkNumbers(qty.value, true)) {
				alert("Enter Numerals Only");
				qty.focus();
				obj.selectedIndex = 0;
				return false;
			}
			ensureScheduleNode(item, "S", "", itemCode, classCode, entryNo);
			optName = getAttributeDisplayName(attrList || "");
			tempValues = qty.value + ":" + itemCode + ":" + classCode + ":" + entryNo + ":" + (field("hUnit") ? field("hUnit").value : "") + ":" + (uom || getAttr(item, "UOM")) + ":" + optName;
			openModal("MRGenSchedulePoP.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:510px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No");
		}
		setAttr(item, "REQUIREDBY", selected);
		return false;
	};

	window.CreateNew = function (obj) {
		var div = field("idConsumption");
		if (div && obj && obj.selectedIndex === 1) {
			div.style.display = "block";
		}
		return false;
	};

	window.hideDiv = function () {
		var div = field("idConsumption");
		if (div) {
			div.style.display = "none";
		}
		if (field("selAccHead")) {
			field("selAccHead").selectedIndex = 0;
		}
		return false;
	};

	window.CheckEntry = function () {
		var desc = field("txtCHead");
		var account = field("selAcc");
		var root = xmlRoot(xmlIsland("OutData"));
		var doc = xmlDocument(xmlIsland("OutData"));
		var node;
		var row;
		if (!desc || !trim(desc.value)) {
			alert("Enter Consumption Head");
			if (desc) {
				desc.select();
			}
			return false;
		}
		if (!account || account.selectedIndex === 0) {
			alert("Select Account Head");
			if (account) {
				account.focus();
			}
			return false;
		}
		node = createElement(doc, "AccountHead");
		setAttr(node, "ACCHEAD", account.value);
		setAttr(node, "CONSUM", trim(desc.value));
		setAttr(node, "ISSFOR", field("hIssueToCode") ? field("hIssueToCode").value : "");
		setAttr(node, "SRC", "N");
		root.appendChild(node);
		if (document.getElementById("tblData")) {
			row = document.getElementById("tblData").insertRow(-1);
			appendCell(row, "ExcelDisplayCell", trim(desc.value));
			appendCell(row, "ExcelDisplayCell", selectedText(account));
		}
		account.selectedIndex = 0;
		desc.value = "";
		return false;
	};

	window.PopDone = function () {
		var root = xmlRoot(xmlIsland("OutData"));
		var select = field("selAccHead");
		syncPost("ConsumptionHeadInsert.asp", serializeXml(xmlIsland("OutData")));
		if (select) {
			elementChildren(root, "AccountHead").forEach(function (node) {
				if (getAttr(node, "SRC") === "N") {
					select.options[select.options.length] = new Option(getAttr(node, "CONSUM"), getAttr(node, "ACCHEAD"));
				}
			});
		}
		window.hideDiv();
		return false;
	};

	window.CheckSubmit = function (todaysDate) {
		var remarks = field("txtRemarks") ? trim(field("txtRemarks").value) : "";
		var action;
		var xhr;
		var message = "";
		setIssueToFromSelection();
		if (field("selIssueTo") && field("selIssueTo").selectedIndex <= 0) {
			alert("Select Requested By");
			field("selIssueTo").focus();
			return false;
		}
		if (selectedIssueType() === "SEL" || selectedIssueType() === "") {
			alert("Select Requested For");
			if (field("cmbIssType")) {
				field("cmbIssType").focus();
			}
			return false;
		}
		if (dateDiffDays(todaysDate, datePickerValue()) > 0) {
			alert("Created On should be less than or equal to Today's Date");
			return false;
		}
		if (remarks.length > 200) {
			alert("Remarks should be less than 200 characters");
			field("txtRemarks").select();
			return false;
		}
		if (!validateEntryItems()) {
			return false;
		}
		removeHeaderNodes();
		if (isCreateMode()) {
			addCreateHeader();
			syncPost("XMLSave.asp?Name=MRS&SessionFlag=true", serializeXml(detailsIsland()));
			form().action = "MRGenerationInsert.asp";
			form().submit();
			return false;
		}
		addApprovalHeader();
		window.DeleteItems();
		action = field("hAction") ? field("hAction").value : "";
		xhr = syncPost("MRApprovalInsert.asp?Action=" + encodeURIComponent(action), serializeXml(detailsIsland()));
		if (trim(xhr.responseText) === "") {
			if (trim(action) === "Amend") {
				message = "Material Requisition has been Updated. Do you want to approve another one?";
			} else if (trim(action) === "Approve") {
				message = "Material Requisition has been Approved / Rejected. Do you want to approve another one?";
			} else if (trim(action) === "Cancel") {
				message = "Material Requisition has been Cancelled. Do you want to approve another one?";
			}
			window.location.href = confirm(message) ? "MRSMGMTLIST.ASP" : "../welcome_Inventory.asp";
		} else {
			alert(xhr.responseText);
		}
		return false;
	};

	window.checkNumbers = checkNumbers;
}(window, document));
