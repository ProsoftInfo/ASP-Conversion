(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
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

	function checkNumbers(value) {
		return /^([0-9]+(\.[0-9]*)?|\.[0-9]+)$/.test(trim(value));
	}

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var year;
		var date;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			year = Number(match[3]);
			if (year < 100) {
				year += 2000;
			}
			date = new Date(year, Number(match[2]) - 1, Number(match[1]));
			return date.getFullYear() === year && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[1]) ? date : null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
		if (match) {
			date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
			return date.getFullYear() === Number(match[1]) && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[3]) ? date : null;
		}
		return null;
	}

	function formatDate(value) {
		var date = parseDate(value);
		return date ? pad2(date.getDate()) + "/" + pad2(date.getMonth() + 1) + "/" + date.getFullYear() : trim(value);
	}

	function dateDiffDays(start, end) {
		var startDate = parseDate(start);
		var endDate = parseDate(end);
		if (!startDate || !endDate) {
			return 0;
		}
		return Math.floor((endDate.getTime() - startDate.getTime()) / 86400000);
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name);
	}

	function xmlDocument(value) {
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
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function loadXml(islandName, text) {
		var island = xmlIsland(islandName);
		if (island && typeof island.loadXML === "function") {
			return island.loadXML(text);
		}
		return false;
	}

	function serializeXml(value) {
		var doc = value && value.nodeType === 9 ? value : xmlDocument(value);
		return doc ? new XMLSerializer().serializeToString(doc) : "";
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
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function importForDocument(doc, node) {
		if (!doc || !node) {
			return null;
		}
		return doc.importNode ? doc.importNode(node, true) : node.cloneNode(true);
	}

	function createElement(doc, name) {
		return doc.createElement(name);
	}

	function clearChildrenByName(node, name) {
		var children = elementChildren(node, name);
		for (var i = 0; i < children.length; i += 1) {
			node.removeChild(children[i]);
		}
	}

	function resetIslandRoot(islandName, rootName) {
		var doc = xmlDocument(xmlIsland(islandName));
		while (doc && doc.firstChild) {
			doc.removeChild(doc.firstChild);
		}
		if (doc) {
			doc.appendChild(doc.createElement(rootName));
		}
		return xmlRoot(xmlIsland(islandName));
	}

	function selectedValue(name) {
		var control = field(name);
		if (!control) {
			return "";
		}
		if (control.options && control.selectedIndex >= 0) {
			return control.options[control.selectedIndex].value;
		}
		return control.value || "";
	}

	function fieldByPrefix(prefix) {
		var frm = form();
		var controls = frm && frm.elements || [];
		for (var i = 0; i < controls.length; i += 1) {
			if (String(controls[i].name || "").indexOf(prefix) === 0) {
				return controls[i];
			}
		}
		return null;
	}

	function controlValue(control) {
		if (!control) {
			return "";
		}
		if (control.options && control.selectedIndex >= 0) {
			return control.options[control.selectedIndex].value;
		}
		return control.value || "";
	}

	function issueTypeValue() {
		var control = field("selIssType");
		if (!control) {
			return "";
		}
		if (control.options) {
			return selectedValue("selIssType");
		}
		return control.checked ? "M" : "F";
	}

	function checkedValue(name, checkedValue, uncheckedValue) {
		var control = field(name);
		return control && control.checked ? checkedValue : uncheckedValue;
	}

	function islandExists(name) {
		ensureCompat();
		return !!(document.getElementById(name) || window[name] || document[name]);
	}

	function detailsIslandName() {
		return islandExists("OutData2") ? "OutData2" : "OutData";
	}

	function detailsIsland() {
		return xmlIsland(detailsIslandName());
	}

	function detailsRoot() {
		return xmlRoot(detailsIsland());
	}

	function getDatePicker() {
		ensureCompat();
		return field("ctlIssDate");
	}

	function getDatePickerValue() {
		var control = getDatePicker();
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		return formatDate(control.value || "");
	}

	function setDatePickerValue(value) {
		var control = getDatePicker();
		if (!control || !value) {
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

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		window.open(url, "_blank");
		return null;
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
		} catch (ignoreHeader) {}
		xhr.send(body == null ? null : body);
		return xhr;
	}

	function issueQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyPPX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function returnTypeField(itemCode, classCode, entryNo) {
		return field("selReturnZ" + itemCode + "Z" + classCode + "Z" + entryNo);
	}

	function approvedQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyA" + classCode + "A" + itemCode + "A" + entryNo);
	}

	function pendingQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyPenX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function transferQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyTraX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function prQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyPrX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function totalQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyTotX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function selectedIndex(name) {
		var control = field(name);
		return control && typeof control.selectedIndex === "number" ? control.selectedIndex : -1;
	}

	function storeValueFor(classCode, itemCode, entryNo, approvedControl) {
		var frm = form();
		var controls = frm && frm.elements || [];
		var prefix = "txtStockZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z";
		var parts;
		for (var i = 0; i < controls.length; i += 1) {
			if (controls[i] === approvedControl && i >= 2 && /^hStoA/i.test(controls[i - 2].name || "")) {
				return controls[i - 2].value || "";
			}
		}
		for (var j = 0; j < controls.length; j += 1) {
			if (String(controls[j].name || "").indexOf(prefix) === 0) {
				parts = String(controls[j].name).split("Z");
				return (parts[5] || "") + ":" + (parts[6] || "");
			}
		}
		return "";
	}

	function firstRootWithItems() {
		var root1 = xmlRoot(xmlIsland("OutData1"));
		var root2 = xmlRoot(xmlIsland("OutData2"));
		if (root1 && (elementChildren(root1, "ITEMDETAILS").length || elementChildren(root1, "ITEM").length)) {
			return root1;
		}
		if (root2 && (elementChildren(root2, "ITEMDETAILS").length || elementChildren(root2, "ITEM").length)) {
			return root2;
		}
		return root1 || root2;
	}

	function matchingItem(root, entryNo, itemCode, classCode) {
		var nodes = elementChildren(root);
		for (var i = 0; i < nodes.length; i += 1) {
			if ((String(nodes[i].nodeName).toLowerCase() === "itemdetails" || String(nodes[i].nodeName).toLowerCase() === "item") &&
					(getAttr(nodes[i], "ENTRYNO") === trim(entryNo) || getAttr(nodes[i], "ItemEntNo") === trim(entryNo)) &&
					(getAttr(nodes[i], "ITEMCODE") === trim(itemCode) || getAttr(nodes[i], "ITMCODE") === trim(itemCode)) &&
					(getAttr(nodes[i], "CLASSCODE") === trim(classCode) || getAttr(nodes[i], "CLACODE") === trim(classCode))) {
				return nodes[i];
			}
		}
		return null;
	}

	function parseAttributeList(value) {
		var text = trim(value).replace(/#/g, "$");
		var parts = text ? text.split(",") : [];
		var attList = [];
		var attId = [];
		var entry;
		var left;
		var attrParts;
		for (var i = 0; i < parts.length; i += 1) {
			entry = parts[i].split("@");
			left = entry[0] || "";
			attrParts = left.split("$");
			if (entry.length > 1) {
				if (trim(attrParts[1]) && trim(attrParts[1]) !== "0") {
					attList.push(attrParts[0]);
					attId.push(attrParts[1]);
				}
			} else if (trim(left)) {
				attList.push(left);
			}
		}
		return {
			attributeList: attList.join(","),
			attributeId: attId.join(",")
		};
	}

	function updatePickQuantity(itemCode, classCode, entryNo) {
		var root = xmlRoot(xmlIsland("OutData1"));
		var item = matchingItem(root, entryNo, itemCode, classCode);
		var picks = elementChildren(item, "Pick");
		var qty = issueQtyField(classCode, itemCode, entryNo);
		var total;
		if (!item) {
			return;
		}
		if (getAttr(root, "DONE") === "NO") {
			return;
		}
		if (picks.length) {
			total = getAttr(picks[0], "TOT") || "0";
			if (qty) {
				qty.value = total;
				qty.readOnly = toNumber(total) !== 0;
				qty.disabled = false;
				qty.className = toNumber(total) !== 0 ? "FormElemRead" : "FormElem";
			}
			setAttr(item, "ONLYLOT", getAttr(picks[0], "ONLYLOT"));
		} else if (qty) {
			qty.value = "0";
			qty.readOnly = false;
			qty.disabled = false;
			qty.className = "FormElem";
		}
	}

	function appendDefaultPick(itemNode, outDoc, issueQty) {
		var itemCode = getAttr(itemNode, "ITMCODE");
		var classCode = getAttr(itemNode, "CLACODE");
		var receiptNumbering = trim(syncGet("../../Common/GetItemRcptNumbering.asp?ItemCode=" + encodeURIComponent(itemCode)).responseText || "");
		var storeResponse = syncGet("../../Common/GetStoreDetailsForItem.asp?ItemCode=" + encodeURIComponent(itemCode) + "&ClassCode=" + encodeURIComponent(classCode));
		var storeDoc = new DOMParser().parseFromString(storeResponse.responseText || "<Root/>", "application/xml");
		var stores = elementChildren(storeDoc.documentElement);
		var pick;
		var detail;
		if (!stores.length || stores.length > 1) {
			alert("Multiple Stores available please select the Issue Quantity for specify store");
			return false;
		}
		pick = createElement(outDoc, "Pick");
		setAttr(pick, "TOT", issueQty);
		setAttr(pick, "NoofPack", "0");
		detail = createElement(outDoc, receiptNumbering === "N" ? "STORE" : "PICK");
		setAttr(detail, "LOC", getAttr(stores[0], "LocNo"));
		setAttr(detail, "BIN", getAttr(stores[0], "BinNo"));
		setAttr(detail, "LOTNO", "N/A");
		setAttr(detail, "INVRECNO", "");
		setAttr(detail, "QTYISS", issueQty);
		setAttr(detail, "NoofPack", "0");
		pick.appendChild(detail);
		itemNode.appendChild(pick);
		return true;
	}

	function appendImmediateSchedule(itemNode, outDoc, nodeName, selectName, scheduleDate) {
		var select = field(selectName) || fieldByPrefix(selectName);
		var schedule;
		if (!select || select.selectedIndex !== 0) {
			return;
		}
		clearChildrenByName(itemNode, nodeName);
		clearChildrenByName(itemNode, "ScheduleDetails");
		schedule = createElement(outDoc, nodeName);
		setAttr(schedule, "STYPE", controlValue(select) || "ID");
		setAttr(schedule, "SVALUE", scheduleDate || formatDate(new Date()));
		itemNode.appendChild(schedule);
	}

	function selectedText(name) {
		var control = field(name);
		if (!control || !control.options || control.selectedIndex < 0) {
			return "";
		}
		return control.options[control.selectedIndex].text;
	}

	function setDisplayText(id, value) {
		var item = field(id);
		if (!item) {
			return;
		}
		if ("value" in item && !/^(span|div|td)$/i.test(item.tagName || "")) {
			item.value = value == null ? "" : String(value);
		} else {
			item.innerHTML = value == null ? "" : String(value);
		}
	}

	function getXmlText(value) {
		var doc = value && value.nodeType === 9 ? value : xmlDocument(value);
		var root = xmlRoot(value);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function loadXmlIntoIsland(name, value) {
		var xml = getXmlText(value);
		if (xml) {
			loadXml(name, xml);
		}
	}

	function saveIssueDetails() {
		loadXmlIntoIsland("OutData1", xmlIsland("OutData2"));
		syncPost("XMLSave.asp?Name=MRISSUEDETAILS&SessionFlag=true", serializeXml(xmlIsland("OutData1")));
	}

	function attrAny(node, names) {
		for (var i = 0; i < names.length; i += 1) {
			if (trim(getAttr(node, names[i])) !== "") {
				return getAttr(node, names[i]);
			}
		}
		return "";
	}

	function cleanAttributeKey(value) {
		return trim(value).replace(/[#:,\s]/g, "");
	}

	function normalizeAttributeList(value) {
		var text = trim(value);
		return text === "0" ? "" : text;
	}

	function issueDetailsRoot() {
		return xmlRoot(xmlIsland("OutData2")) || resetIslandRoot("OutData2", "root");
	}

	function issueDetailsDoc() {
		return xmlDocument(xmlIsland("OutData2"));
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

	function appendItemDetail(source, options) {
		var root = issueDetailsRoot();
		var doc = issueDetailsDoc();
		var itemCode = attrAny(source, ["ItemCode", "ITEMCODE", "ITMCODE"]);
		var classCode = attrAny(source, ["ClassCode", "CLASSCODE", "CLACODE"]);
		var entryNo = attrAny(source, ["EntryNo", "ENTRYNO", "ItemEntNo"]) || String(elementChildren(root, "ITEMDETAILS").length + 1);
		var item;
		if (!root || !doc || !itemCode || !classCode || itemDetailExists(root, itemCode, classCode, entryNo)) {
			return;
		}
		item = createElement(doc, "ITEMDETAILS");
		setAttr(item, "ENTRYNO", entryNo);
		setAttr(item, "ITEMCODE", itemCode);
		setAttr(item, "CLASSCODE", classCode);
		setAttr(item, "UNIT", attrAny(source, ["UNIT", "OrgCode"]) || (field("hUnit") ? field("hUnit").value : ""));
		setAttr(item, "ITEMNAME", attrAny(source, ["ItemName", "ITEMNAME", "ITMNAME"]).replace(/"/g, "~~"));
		setAttr(item, "UOM", attrAny(source, ["StoresUoM", "UOM", "UoM"]));
		setAttr(item, "DECIMAL", attrAny(source, ["Decimal", "DECIMAL"]) || "0");
		setAttr(item, "DISPALYED", "N");
		setAttr(item, "QTY", attrAny(source, ["Qty", "QTY"]) || "0");
		setAttr(item, "REQUIREDBY", "");
		setAttr(item, "REQUIREDVALUE", "");
		setAttr(item, "ATTRIBUTELIST", normalizeAttributeList(attrAny(source, ["AttributeList", "ATTRIBUTELIST"])));
		setAttr(item, "REMARKS", "");
		setAttr(item, "RefNo", attrAny(source, ["No", "RefNo", "MRSNO"]));
		setAttr(item, "ReqQty", attrAny(source, ["Qty", "ReqQty", "REQQTY"]) || "0");
		setAttr(item, "ONLYLOT", "");
		setAttr(item, "RETURNABLE", attrAny(source, ["RETURNABLE"]) || "N");
		setAttr(item, "RETURNITEM", attrAny(source, ["RETURNITEM"]) || "S");
		root.appendChild(item);
		if (options && options.detectItemType && field("hItemType") && !trim(field("hItemType").value)) {
			setItemTypeForItem(itemCode);
		}
	}

	function setItemTypeForItem(itemCode) {
		var xhr;
		var root;
		if (!itemCode || !field("hItemType") || trim(field("hItemType").value)) {
			return;
		}
		xhr = syncGet("../../Include/GetItemTypeForItem.asp?ItemCode=" + encodeURIComponent(itemCode));
		if (xhr.responseText) {
			loadXml("ItemTypeData", xhr.responseText);
			root = xmlRoot(xmlIsland("ItemTypeData"));
			if (root) {
				field("hItemType").value = getAttr(root, "ItemType");
			}
		}
	}

	function processReferenceHeader(root, options) {
		var refs = elementChildren(root, "Reference");
		var refCodes = [];
		var refNos = [];
		var refDates = [];
		var ref;
		var remarks;
		var partyParts;
		for (var i = 0; i < refs.length; i += 1) {
			ref = refs[i];
			refCodes.push(getAttr(ref, "ReferenceCode") + " - " + getAttr(ref, "ReferenceDate"));
			refNos.push(getAttr(ref, "ReferenceNo"));
			refDates.push(getAttr(ref, "ReferenceDate"));
			if (options && options.material && trim(field("hType") && field("hType").value) === "SUB") {
				remarks = getAttr(ref, "Remarks");
				partyParts = remarks.split("-");
				if (partyParts.length > 1) {
					setDisplayText("txtParty", partyParts[1]);
					if (field("hIssueToType")) {
						field("hIssueToType").value = "Party";
					}
					if (field("hIssueToCode")) {
						field("hIssueToCode").value = partyParts[0];
					}
					selectOptionByValue("selIssueTo", "Party");
				}
			}
		}
		if (refNos.length) {
			setDisplayText("RefNoDate", refCodes.join(","));
			if (field("hRefNo")) {
				field("hRefNo").value = refNos.join(",");
			}
			if (field("hRefDate")) {
				field("hRefDate").value = refDates.join(",");
			}
		}
		return refNos.join(",");
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

	function processItemSelection(selectionRoot, options) {
		var root = selectionRoot || xmlRoot(xmlIsland("OutData"));
		var refType = options && options.refType || selectedValue("selRefName");
		var refNos;
		var xhr;
		var itemRoot;
		if (!root) {
			return;
		}
		loadXmlIntoIsland("OutData", root);
		if (trim(refType) !== "N") {
			refNos = processReferenceHeader(root, options);
			if (refNos) {
				xhr = syncGet("InvGetItemDetForRefType.asp?RefType=" + encodeURIComponent(refType) + "&RefCodes=" + encodeURIComponent(refNos) + "&orgID=" + encodeURIComponent(field("hUnit") ? field("hUnit").value : ""));
				if (xhr.responseText) {
					loadXml("ItemData", xhr.responseText);
					itemRoot = xmlRoot(xmlIsland("ItemData"));
					elementChildren(itemRoot).forEach(function (node) {
						if (/^item$/i.test(node.nodeName)) {
							appendItemDetail(node, { detectItemType: true });
						} else if (/^subcontract$/i.test(node.nodeName)) {
							issueDetailsRoot().appendChild(importForDocument(issueDetailsDoc(), node));
						}
					});
				}
			}
		} else {
			elementChildren(root).forEach(function (node) {
				if (/^item$/i.test(node.nodeName)) {
					appendItemDetail(node, { detectItemType: true });
				}
			});
		}
		window.DisplayTable(new Date());
		saveIssueDetails();
	}

	function stockFor(itemCode, classCode, entryNo, attributeList) {
		var para = itemCode + ":" + entryNo + ":" + classCode + ":" + (field("hUnit") ? field("hUnit").value : field("hOrgID") ? field("hOrgID").value : "") + ":" + (field("hMinDate") ? field("hMinDate").value : "") + ":" + (field("hMaxDate") ? field("hMaxDate").value : "") + ":" + (attributeList || "");
		var xhr = syncGet("XMLGetStockDetails.asp?Para=" + encodeURIComponent(para));
		var doc;
		var root;
		var nodes;
		if (!xhr.responseText) {
			return "0";
		}
		doc = new DOMParser().parseFromString(xhr.responseText, "application/xml");
		root = doc.documentElement;
		nodes = elementChildren(root);
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "ITEMCODE") === trim(itemCode) && getAttr(nodes[i], "CLASSCODE") === trim(classCode)) {
				return getAttr(nodes[i], "STOCK") || "0";
			}
		}
		return "0";
	}

	function appendHidden(container, name, value) {
		var input = document.createElement("input");
		input.type = "hidden";
		input.name = name;
		input.value = value == null ? "" : String(value);
		container.appendChild(input);
		return input;
	}

	function appendTextInput(container, name, value, size, readOnly) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.value = value == null ? "" : String(value);
		input.size = size || 10;
		input.className = readOnly ? "FormElemRead" : "FormElem";
		input.style.textAlign = "right";
		input.readOnly = !!readOnly;
		container.appendChild(input);
		return input;
	}

	function addCell(row, className, text) {
		var cell = row.insertCell(-1);
		cell.className = className || "ExcelDisplayCell";
		if (text != null) {
			cell.innerHTML = text;
		}
		return cell;
	}

	function makeIcon(src, alt, handler) {
		var link = document.createElement("a");
		var img = document.createElement("img");
		link.href = "#";
		link.onclick = function () {
			handler(img);
			return false;
		};
		img.border = "0";
		img.src = src;
		img.width = 15;
		img.height = 15;
		img.alt = alt;
		link.appendChild(img);
		return { link: link, img: img };
	}

	function clearRenderedRows(table, headerRows) {
		while (table && table.rows.length > headerRows) {
			table.deleteRow(headerRows);
		}
	}

	function renderDirectRow(table, item, index) {
		var itemCode = getAttr(item, "ITEMCODE");
		var classCode = getAttr(item, "CLASSCODE");
		var entryNo = getAttr(item, "ENTRYNO");
		var attrList = getAttr(item, "ATTRIBUTELIST");
		var itemName = getAttr(item, "ITEMNAME").replace(/~~/g, "\"");
		var stock = stockFor(itemCode, classCode, entryNo, attrList);
		var row = table.insertRow(-1);
		var detailRow = table.insertRow(-1);
		var cell;
		var icon;
		addCell(row, "ExcelDisplayCell", index).rowSpan = 2;
		cell = addCell(row, "ExcelDisplayCell");
		cell.rowSpan = 2;
		cell.innerHTML = "<a class='ExcelDisplayLink' href='#' name='lnkA" + itemCode + "A" + classCode + "A" + getAttr(item, "UNIT") + "'>" + itemName + "</a>";
		cell.firstChild.onclick = function () {
			window.DisplayItem(this.name);
			return false;
		};
		addCell(row, "ExcelDisplayCell").appendChild(appendTextInput(document.createElement("span"), "", "", 1, true));
		row.cells[row.cells.length - 1].innerHTML = stock;
		addCell(row, "ExcelDisplayCell", "0");
		addCell(row, "ExcelDisplayCell", "0");
		cell = addCell(row, "ExcelInputCell");
		icon = makeIcon("../../assets/images/iTMS%20Icons/Entry.gif", "Pick Details", function (img) {
			img.name = "btnZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z" + "" + "Z" + getAttr(item, "UNIT") + "Z" + attrList;
			window.CheckLot(img, "", attrList);
		});
		cell.appendChild(icon.link);
		appendHidden(cell, "hStoName" + classCode + "A" + itemCode + "A" + entryNo, "0");
		appendHidden(cell, "hSchA" + classCode + "A" + itemCode + "A" + entryNo, "");
		appendHidden(cell, "hStoA0A0", "0");
		appendTextInput(cell, "txtQtyPPX" + classCode + "X" + itemCode + "X" + entryNo, getAttr(item, "QTY"), 10, true);
		addCell(row, "ExcelDisplayCell", "");
		cell = addCell(row, "ExcelDisplayCell");
		cell.rowSpan = 2;
		cell.appendChild(addAdditionalButton(classCode, itemCode, getAttr(item, "UNIT"), entryNo));
		addCell(detailRow, "ExcelDisplayCell", "0");
		addCell(detailRow, "ExcelDisplayCell", "0");
		addCell(detailRow, "ExcelDisplayCell", "0");
		addCell(detailRow, "ExcelDisplayCell", "");
		addCell(detailRow, "ExcelDisplayCell", getAttr(item, "UOM"));
	}

	function addAdditionalButton(classCode, itemCode, orgId, entryNo) {
		var button = document.createElement("input");
		button.type = "button";
		button.name = "btnAddDet" + classCode + "A" + itemCode + "A" + entryNo;
		button.className = "ActionButtonX";
		button.value = "Yes";
		button.onclick = function () {
			window.GetAddDet(classCode, itemCode, orgId, entryNo);
			return false;
		};
		return button;
	}

	function renderMaterialRow(table, item, index) {
		var itemCode = getAttr(item, "ITEMCODE");
		var classCode = getAttr(item, "CLASSCODE");
		var entryNo = getAttr(item, "ENTRYNO");
		var attrList = getAttr(item, "ATTRIBUTELIST");
		var attrKey = cleanAttributeKey(attrList);
		var itemName = getAttr(item, "ITEMNAME").replace(/~~/g, "\"");
		var reqQty = getAttr(item, "ReqQty") || getAttr(item, "QTY") || "0";
		var stock = stockFor(itemCode, classCode, entryNo, attrList);
		var row = table.insertRow(-1);
		var cell;
		var checkbox;
		var select;
		var icon;
		addCell(row, "ExcelDisplayCell", index);
		cell = addCell(row, "ExcelInputCell");
		checkbox = document.createElement("input");
		checkbox.type = "checkbox";
		checkbox.name = "chkItemZ" + itemCode + "Z" + classCode + "Z" + attrKey;
		checkbox.className = "FormElem";
		cell.appendChild(checkbox);
		cell = addCell(row, "ExcelDisplayCell");
		cell.innerHTML = "<a class='ExcelDisplayLink' href='#' name='lnkA" + itemCode + "A" + classCode + "A" + getAttr(item, "UNIT") + "'>" + itemName + "</a>";
		cell.firstChild.onclick = function () {
			window.DisplayItem(this.name);
			return false;
		};
		cell = addCell(row, "ExcelInputCell");
		select = document.createElement("select");
		select.name = "selReturnZ" + itemCode + "Z" + classCode + "Z" + entryNo;
		select.id = select.name;
		select.className = "FormElem";
		[["N", "No"], ["Y", "Yes"], ["D", "Diff."]].forEach(function (entry) {
			var option = document.createElement("option");
			option.value = entry[0];
			option.text = entry[1];
			select.options.add(option);
		});
		cell.appendChild(select);
		cell = addCell(row, "ExcelDisplayCell", reqQty);
		appendHidden(cell, "txtRequestedZ" + itemCode + "Z" + classCode + "Z" + entryNo, reqQty);
		appendHidden(cell, "txtToIssueZ" + itemCode + "Z" + classCode + "Z" + entryNo, "0");
		addCell(row, "ExcelDisplayCell", stock);
		cell = addCell(row, "ExcelInputCell");
		icon = makeIcon("../../assets/images/iTMS%20Icons/Entry.gif", "Pick Details", function (img) {
			img.name = "btnZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z" + "" + "Z" + getAttr(item, "UNIT") + "Z" + attrList;
			window.CheckLot(img, "", attrList.replace(/:/g, "@"));
		});
		cell.appendChild(icon.link);
		appendHidden(cell, "hStoName" + classCode + "A" + itemCode + "A" + entryNo, "0");
		appendHidden(cell, "hSchA" + classCode + "A" + itemCode + "A" + entryNo, "");
		appendHidden(cell, "hStoA0A0", "0");
		appendHidden(cell, "hRcptNumA" + classCode + "A" + itemCode + "A" + entryNo, "N");
		appendTextInput(cell, "txtQtyPPX" + classCode + "X" + itemCode + "X" + entryNo, getAttr(item, "QTY"), 10, false);
		cell = addCell(row, "ExcelInputCell");
		icon = makeIcon("../../assets/images/iTMS%20Icons/Entry.gif", "Pick Schedule", function (img) {
			img.name = "btnScheduleZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z" + "" + "Z" + getAttr(item, "UNIT") + "Z" + attrList;
			window.ScheduleDate(img, "", attrList.replace(/:/g, "@"));
		});
		cell.appendChild(icon.link);
		addCell(row, "ExcelDisplayCell", "");
		cell = addCell(row, "ExcelDisplayCell");
		cell.appendChild(addAdditionalButton(classCode, itemCode, getAttr(item, "UNIT"), entryNo));
	}

	function processDetailsForDisplay() {
		var root = issueDetailsRoot();
		return elementChildren(root, "ITEMDETAILS");
	}

	function buildIssueXml(todaysDate) {
		var sourceRoot = firstRootWithItems();
		var sourceItems = elementChildren(sourceRoot);
		var targetRoot = resetIslandRoot("OutData1", "ISSTYPE");
		var targetDoc = xmlDocument(xmlIsland("OutData1"));
		var retCount = 0;
		var issueDate = getDatePickerValue();
		var refType = selectedValue("selRefName");
		var issueMode = field("hIssMode") ? field("hIssMode").value : "N";
		var parsedAttrs;
		var itemNode;
		var itemCode;
		var classCode;
		var entryNo;
		var itemName;
		var issueQty;
		var transferQty;
		var prQty;
		var totalQty;
		var approvedQty;
		var qtyControl;
		var approvedControl;
		var pendingControl;
		var transferControl;
		var prControl;
		var totalControl;
		var isMRSItem;
		var nodeName;
		var retControl;
		var retValue;
		var converted;
		var children;
		var child;
		var auxItem;
		var auxChildren;

		if (dateDiffDays(todaysDate, issueDate) > 0) {
			alert("Issue Date should be less than or equal to Today's Date");
			return null;
		}
		if (field("hAutoConsumption") && field("hAutoConsumption").value === "Y") {
			if (field("selAccHead") && field("selAccHead").selectedIndex <= 1) {
				alert("Select the Account Head");
				field("selAccHead").focus();
				return null;
			}
		}
		if (trim(field("hType") && field("hType").value) === "SUB" && !trim(field("hIssueToCode") && field("hIssueToCode").value)) {
			alert("Please Select Issued To");
			return null;
		}

		for (var i = 0; i < sourceItems.length; i += 1) {
			itemNode = sourceItems[i];
			if (String(itemNode.nodeName).toLowerCase() === "subcontract") {
				targetRoot.appendChild(importForDocument(targetDoc, itemNode));
				continue;
			}
			if (String(itemNode.nodeName).toLowerCase() !== "itemdetails" && String(itemNode.nodeName).toLowerCase() !== "item") {
				continue;
			}
			entryNo = getAttr(itemNode, "ENTRYNO");
			itemCode = getAttr(itemNode, "ITEMCODE") || getAttr(itemNode, "ITMCODE");
			classCode = getAttr(itemNode, "CLASSCODE") || getAttr(itemNode, "CLACODE");
			itemName = getAttr(itemNode, "ITEMNAME") || getAttr(itemNode, "ITMNAME");
			nodeName = String(itemNode.nodeName).toLowerCase();
			isMRSItem = nodeName === "item" && !!getAttr(itemNode, "MRSNO");
			qtyControl = issueQtyField(classCode, itemCode, entryNo);
			approvedControl = approvedQtyField(classCode, itemCode, entryNo);
			pendingControl = pendingQtyField(classCode, itemCode, entryNo);
			transferControl = transferQtyField(classCode, itemCode, entryNo);
			prControl = prQtyField(classCode, itemCode, entryNo);
			totalControl = totalQtyField(classCode, itemCode, entryNo);
			issueQty = trim(qtyControl && qtyControl.value) || "0";
			transferQty = trim(transferControl && transferControl.value) || "0";
			prQty = trim(prControl && prControl.value) || "0";
			approvedQty = trim(approvedControl && approvedControl.value) || getAttr(itemNode, "REQQTY") || "0";
			totalQty = toNumber(issueQty) + toNumber(transferQty) + toNumber(prQty);
			if (!checkNumbers(issueQty) || !checkNumbers(transferQty) || !checkNumbers(prQty)) {
				alert("Enter Numerals Only");
				if (qtyControl && !checkNumbers(issueQty)) {
					qtyControl.focus();
				} else if (transferControl && !checkNumbers(transferQty)) {
					transferControl.focus();
				} else if (prControl && !checkNumbers(prQty)) {
					prControl.focus();
				}
				return null;
			}
			if (totalControl) {
				totalControl.value = String(totalQty);
			}
			if (isMRSItem) {
				if (pendingControl && totalQty > toNumber(pendingControl.value)) {
					alert("Total Quantity should be less than or equal to Quantity Pending (" + pendingControl.value + ")");
					if (qtyControl) {
						qtyControl.focus();
					}
					return null;
				}
			} else {
				if (toNumber(issueQty) <= 0) {
					alert("Issue Quantity Should be Greater then Zero for " + itemName);
					return null;
				}
			}
			if (refType === "11" && issueMode !== "E") {
				if (toNumber(issueQty) > toNumber(field("txtRequestedZ" + itemCode + "Z" + classCode + "Z" + entryNo) && field("txtRequestedZ" + itemCode + "Z" + classCode + "Z" + entryNo).value) - toNumber(field("txtToIssueZ" + itemCode + "Z" + classCode + "Z" + entryNo) && field("txtToIssueZ" + itemCode + "Z" + classCode + "Z" + entryNo).value)) {
					alert("Issue Quantity Must be Less than or Equal to Request Quantity");
					return null;
				}
			}

			parsedAttrs = parseAttributeList(getAttr(itemNode, "ATTRIBUTELIST"));
			converted = createElement(targetDoc, "ITEM");
			setAttr(converted, "ENTRYNO", entryNo);
			setAttr(converted, "ITMCODE", itemCode);
			setAttr(converted, "CLACODE", classCode);
			setAttr(converted, "ITMNAME", itemName);
			setAttr(converted, "SSTORE", isMRSItem ? storeValueFor(classCode, itemCode, entryNo, approvedControl) : "");
			setAttr(converted, "REQQTY", isMRSItem ? approvedQty : "0");
			setAttr(converted, "REQBY", field("txtRecBy") ? field("txtRecBy").value : "");
			setAttr(converted, "REMARKS", field("Remarks") ? field("Remarks").value : "");
			setAttr(converted, "ITEMTYPE", field("hItemType") ? field("hItemType").value : field("hItmType") ? field("hItmType").value : "");
			setAttr(converted, "ISSUEDATE", issueDate);
			setAttr(converted, "ISSQTY", issueQty);
			setAttr(converted, "TRAQTY", transferQty);
			setAttr(converted, "PRQTY", prQty);
			setAttr(converted, "IVALUE", isMRSItem ? getAttr(itemNode, "IVALUE") || "0" : issueQty);
			setAttr(converted, "ORGCODE", getAttr(itemNode, "UNIT") || field("hOrgID") && field("hOrgID").value || field("hUnit") && field("hUnit").value || "");
			setAttr(converted, "MRSNO", getAttr(itemNode, "MRSNO") || (refType === "11" && field("hRefNo") ? field("hRefNo").value : ""));
			setAttr(converted, "MRSDATE", getAttr(itemNode, "MRSDATE") || (refType === "11" && field("hRefDate") ? field("hRefDate").value : ""));
			setAttr(converted, "ATTRIBUTELIST", isMRSItem ? getAttr(itemNode, "ATTRIBUTELIST") : parsedAttrs.attributeId || parsedAttrs.attributeList);
			setAttr(converted, "CREATEDBY", getAttr(itemNode, "CREATEDBY") || (field("hUserId") ? field("hUserId").value : ""));
			setAttr(converted, "CREATEDON", getAttr(itemNode, "CREATEDON") || issueDate);
			setAttr(converted, "RefNo", getAttr(itemNode, "RefNo") || getAttr(itemNode, "MRSNO"));
			setAttr(converted, "ONLYLOT", getAttr(itemNode, "ONLYLOT"));

			retControl = returnTypeField(itemCode, classCode, entryNo);
			retValue = selectedValue("selReturnZ" + itemCode + "Z" + classCode + "Z" + entryNo) || "N";
			if (retControl && retControl.selectedIndex > 0) {
				retCount += 1;
			}
			if (retValue === "N") {
				setAttr(converted, "RETURNABLE", "N");
				setAttr(converted, "RETURNITEM", "S");
			} else if (retValue === "Y") {
				setAttr(converted, "RETURNABLE", "Y");
				setAttr(converted, "RETURNITEM", "S");
			} else {
				setAttr(converted, "RETURNABLE", "Y");
				setAttr(converted, "RETURNITEM", "D");
			}
			setAttr(converted, "MatType", "");

			children = elementChildren(itemNode);
			for (var c = 0; c < children.length; c += 1) {
				child = children[c];
				if (/^(pick|adddet|stschedule|prschedule|scheduledetails)$/i.test(child.nodeName)) {
					converted.appendChild(importForDocument(targetDoc, child));
				}
			}
			auxItem = matchingItem(detailsRoot(), entryNo, itemCode, classCode);
			auxChildren = elementChildren(auxItem);
			for (var ac = 0; ac < auxChildren.length; ac += 1) {
				child = auxChildren[ac];
				if (/^(stschedule|prschedule|scheduledetails)$/i.test(child.nodeName)) {
					converted.appendChild(importForDocument(targetDoc, child));
				}
			}
			if (isMRSItem && toNumber(transferQty) > 0) {
				appendImmediateSchedule(converted, targetDoc, "STSchedule", "selSTSchZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z", todaysDate);
			}
			if (isMRSItem && toNumber(prQty) > 0) {
				appendImmediateSchedule(converted, targetDoc, "PRSchedule", "selPRSchZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z", todaysDate);
			}
			if (toNumber(issueQty) > 0 && !elementChildren(converted, "Pick").length && !appendDefaultPick(converted, targetDoc, issueQty)) {
				return null;
			}
			targetRoot.appendChild(converted);
		}

		if (trim(field("hType") && field("hType").value) === "SUB" && retCount <= 0) {
			alert("SubContract Issue Should have minimum One Returnable Item");
			return null;
		}
		if (!elementChildren(targetRoot, "ITEM").length) {
			alert("No Items are available for Issue");
			return null;
		}
		return targetRoot;
	}

	function decorateIssueRoot(root) {
		var issueType = issueTypeValue();
		var refType = selectedValue("selRefName");
		setAttr(root, "ISSTYPE", issueType);
		setAttr(root, "ISSTOTYPE", field("hIssueToType") ? field("hIssueToType").value : "");
		setAttr(root, "ISSTOCODE", field("hIssueToCode") ? field("hIssueToCode").value : "");
		setAttr(root, "ISSTOSUBCODE", field("hIssueToSubCode") ? field("hIssueToSubCode").value : "");
		setAttr(root, "ISSFORCODE", field("hUsage") ? field("hUsage").value : "");
		setAttr(root, "ISSFORTYPE", field("hIssForType") ? field("hIssForType").value : "");
		setAttr(root, "PARTYCODE", field("hPartyCode") ? field("hPartyCode").value : "");
		setAttr(root, "POConfirm", "N");
		setAttr(root, "SInvConfirm", "N");
		setAttr(root, "Invoice", "A");
		setAttr(root, "GPConfirm", "N");
		setAttr(root, "ProConfirm", "N");
		setAttr(root, "MCallFrom", "MRIssue");
		setAttr(root, "RedirectTo", "ISSUEMGMT.ASP");
		setAttr(root, "MRNo", "");
		setAttr(root, "AppRefType", refType && refType !== "N" ? refType : "");
		setAttr(root, "AppRefNo", field("hRefNo") ? field("hRefNo").value : "");
		setAttr(root, "AppRefDate", field("hRefDate") ? field("hRefDate").value : "");
		setAttr(root, "ConsumptionAccHead", selectedValue("selAccHead"));
		setAttr(root, "IssueToCode", field("hIssueToCode") ? field("hIssueToCode").value : "");
		setAttr(root, "PickPackFlag", field("hPickPackFlag") ? field("hPickPackFlag").value : "");
		setAttr(root, "IssFrom", field("hIssFrom") ? field("hIssFrom").value : "");
		setAttr(root, "Returnable", "N");
		setAttr(root, "ReturnItem", "S");
		setAttr(root, "TYPE", field("hType") ? field("hType").value : "");
	}

	window.Help = function () {
		window.open("../HelpFiles/Issue.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px,left=10,top=10");
		return false;
	};

	window.CheckAvailability = function (itemCode) {
		var root = issueDetailsRoot();
		var nodes = elementChildren(root, "ITEMDETAILS");
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "ITEMCODE") === trim(itemCode)) {
				return true;
			}
		}
		return false;
	};

	window.ClearTable = function () {
		var table = document.getElementById("tblLot");
		if (table) {
			clearRenderedRows(table, field("selIssueTo") ? 2 : 3);
		}
		return false;
	};

	window.DisplayTable = function () {
		var table = document.getElementById("tblLot");
		var items = processDetailsForDisplay();
		var material = !!field("selIssueTo");
		if (!table) {
			return false;
		}
		clearRenderedRows(table, material ? 2 : 3);
		for (var i = 0; i < items.length; i += 1) {
			setAttr(items[i], "DISPALYED", "Y");
			if (material) {
				renderMaterialRow(table, items[i], i + 1);
			} else {
				renderDirectRow(table, items[i], i + 1);
			}
		}
		if (field("hCtr")) {
			field("hCtr").value = String(table.rows.length);
		}
		window.PackData();
		return false;
	};

	window.DeleteItem = function () {
		var root = issueDetailsRoot();
		var nodes = elementChildren(root, "ITEMDETAILS");
		var node;
		var checkbox;
		for (var i = nodes.length - 1; i >= 0; i -= 1) {
			node = nodes[i];
			checkbox = field("chkItemZ" + getAttr(node, "ITEMCODE") + "Z" + getAttr(node, "CLASSCODE") + "Z" + cleanAttributeKey(getAttr(node, "ATTRIBUTELIST")));
			if (checkbox && checkbox.checked) {
				root.removeChild(node);
			}
		}
		window.DisplayTable(new Date());
		saveIssueDetails();
		return false;
	};

	window.AddItem = function () {
		return window.GetItems(new Date());
	};

	window.ScheduleDate = function (obj, issueQty, attributeList) {
		var parts = String(obj && obj.name || "").split("Z");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var optionName = parts[4] || "";
		var orgId = parts[5] || "";
		var attrs = parseAttributeList(attributeList || parts[6] || "");
		var packFlag = field("hPickPackFlag") ? field("hPickPackFlag").value : "";
		var tempValues = itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + entryNo + ":" + optionName + ":" + entryNo + ":" + optionName + ":" + (issueQty || "") + ":" + issueTypeValue() + ":" + (field("hIssueToCode") ? field("hIssueToCode").value : field("hUsage") ? field("hUsage").value : "") + ":" + orgId + ":" + attrs.attributeList + ":" + attrs.attributeId;
		openModal("mrsPickSchedulePoP.asp?sTemp=" + encodeURIComponent(tempValues) + "&AttributeList=" + encodeURIComponent(attributeList || "") + "&PickPackFlag=" + encodeURIComponent(packFlag), xmlIsland("OutData1"), "dialogHeight:350px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No");
		return false;
	};

	window.ReturnData = function (itemCode, classCode, entryNo) {
		var checkbox = field("chkReturnable" + itemCode + "Z" + classCode + "Z" + entryNo);
		var block = field("divItemData" + itemCode + "Z" + classCode + "Z" + entryNo);
		if (block) {
			block.style.display = checkbox && checkbox.checked ? "block" : "none";
		}
		return false;
	};

	window.RetItemData = function () {
		var checked = document.querySelector('input[name="radItem"]:checked');
		window.sReturnItem = checked ? checked.value : "";
		return window.sReturnItem;
	};

	window.MarkData = function () {
		var marked = field("selIssType");
		var div = field("divPickPack");
		if (div) {
			div.style.display = marked && marked.checked ? "block" : "none";
		}
		window.PackData();
		return false;
	};

	window.PackData = function () {
		var pick = document.querySelector('input[name="radPick"]:checked');
		var marked = field("selIssType");
		var inputs;
		var readOnly;
		if (field("hPickPackFlag")) {
			field("hPickPackFlag").value = pick ? pick.value : "";
		}
		inputs = document.querySelectorAll('input[name^="txtQtyPPX"]');
		for (var i = 0; i < inputs.length; i += 1) {
			readOnly = !!(marked && marked.checked && field("hPickPackFlag") && field("hPickPackFlag").value !== "L");
			inputs[i].readOnly = readOnly;
			inputs[i].disabled = false;
			inputs[i].className = readOnly ? "FormElemRead" : "FormElem";
		}
		return false;
	};

	window.GetDetails = function () {
		var refType = selectedValue("selRefName");
		var orgId = field("hUnit") ? field("hUnit").value : "";
		var partyCode = "";
		var issueTo = selectedValue("selIssueTo");
		var material = !!field("selIssueTo");
		if (material && trim(field("hType") && field("hType").value) !== "SUB" && selectedIndex("selIssueTo") === 0) {
			alert("Select Issue To");
			field("selIssueTo").focus();
			return false;
		}
		if (!material) {
			if (getAttr(xmlRoot(xmlIsland("RefData")), "Usage") === "PRD" && getAttr(xmlRoot(xmlIsland("RefData")), "IssueTo") === "M" && refType !== "14") {
				alert("Select Mix Code Reference");
				if (field("selRefName")) {
					field("selRefName").focus();
				}
				return false;
			}
		} else {
			if (field("hIssFrom")) {
				field("hIssFrom").value = selectedValue("cmbIssFrom");
			}
			if (trim(issueTo).toLowerCase() === "party") {
				partyCode = field("hIssueToCode") ? field("hIssueToCode").value : "";
			}
		}
		window.sUnit = orgId;
		window.sIType = field("hItemType") ? field("hItemType").value : "";
		if (typeof window.RefTypeSelection === "function") {
			window.RefTypeSelection(refType, orgId, partyCode, "Y", 1, "N", 0, "", function (result) {
				processItemSelection(xmlRoot(result) || xmlRoot(xmlIsland("OutData")), { refType: refType, material: material });
			});
		}
		return false;
	};

	window.GetItems = function () {
		var orgId = field("hUnit") ? field("hUnit").value : "";
		var usage = field("hIssueToCode") ? field("hIssueToCode").value : field("hUsage") ? field("hUsage").value : "";
		var refCodes = field("hRefCodes") ? field("hRefCodes").value : "";
		var url;
		window.sUnit = orgId;
		window.sIType = field("hItemType") ? field("hItemType").value : "";
		if (trim(usage) === "PRD") {
			url = "../../Include/ItemSelectRefBased.asp?orgID=" + encodeURIComponent(orgId) + "&sIType=FIB&Stock=Y&hSelectMode=M&Flag=1&RefCodes=" + encodeURIComponent(refCodes);
			openModal(url, xmlIsland("OutSelectData"), "dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function (result) {
				processItemSelection(xmlRoot(result) || xmlRoot(xmlIsland("OutSelectData")), { refType: "N", material: !!field("selIssueTo") });
			});
			return false;
		}
		if (typeof window.RefTypeSelection === "function") {
			window.RefTypeSelection("N", orgId, "", "Y", 1, "N", 0, "", function (result) {
				processItemSelection(xmlRoot(result) || xmlRoot(xmlIsland("OutData")), { refType: "N", material: !!field("selIssueTo") });
			});
		}
		return false;
	};

	window.GetAddDet = function (classCode, itemCode, orgId, entryNo) {
		var qty = issueQtyField(classCode, itemCode, entryNo);
		var returnType = field("selReturnZ" + itemCode + "Z" + classCode + "Z" + entryNo);
		var issueType = field("hType") ? field("hType").value : "";
		var usage = field("hIssueToCode") ? field("hIssueToCode").value : field("hUsage") ? field("hUsage").value : "";
		var returnable = "N";
		var returnItem = "S";
		var tempValues;
		if (!qty || trim(qty.value) === "") {
			alert("Enter Quantity");
			return false;
		}
		if (toNumber(qty.value) === 0) {
			return false;
		}
		if (returnType) {
			if (returnType.selectedIndex === 1) {
				returnable = "Y";
			} else if (returnType.selectedIndex === 2) {
				returnable = "Y";
				returnItem = "D";
			}
		}
		if (trim(issueType) === "JWK") {
			if (returnType && returnType.selectedIndex <= 0) {
				alert("Please select the Returnable");
				return false;
			}
			tempValues = classCode + "|" + itemCode + "|" + orgId + "|" + usage + "|" + trim(qty.value) + "||" + returnable + "|" + returnItem + "|" + entryNo;
			openModal("IssSubConProcessDetailsPop.asp?sTemp=" + encodeURIComponent(tempValues), xmlIsland("OutData1"), "dialogHeight:370px;dialogWidth:500px;center:Yes;help:No;resizable:No;status:No");
		} else {
			tempValues = classCode + "|" + itemCode + "|" + orgId + "|" + usage + "|" + trim(qty.value);
			openModal("DirectIssueAddEntry.asp?sTemp=" + encodeURIComponent(tempValues), xmlIsland("OutData2"), "dialogHeight:370px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No");
		}
		return false;
	};

	window.DoChanges = function (obj) {
		var option = obj && obj.options && obj.selectedIndex >= 0 ? obj.options[obj.selectedIndex] : null;
		var value = controlValue(obj);
		var orgId = field("hUnit") ? field("hUnit").value : "";
		var partyFeatures = "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No";
		var selectParty;
		setDisplayText("idUsage", option ? option.text : value);
		if (field("hUsage")) {
			field("hUsage").value = value;
		}
		if (value === "JWK") {
			openModal("JobworkPop.asp?sUnit=" + encodeURIComponent(orgId), "", "dialogHeight:200px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", function (outValue) {
				if (field("hJobWorkNo")) {
					field("hJobWorkNo").value = outValue || "";
				}
			});
		} else if (/^(DIS|IUT|PUR|REP|SAL|SER)$/.test(value)) {
			selectParty = function (url) {
				openModal(url, "", partyFeatures, function (outValue) {
					var parts = String(outValue || "").split(":");
					if (outValue && parts.length === 1) {
						selectParty("SalesInvoicePartyPopup.asp?" + outValue);
						return;
					}
					if (field("hSupplier")) {
						field("hSupplier").value = parts[1] || "";
					}
				});
			};
			selectParty("SalesInvoicePartyPopup.asp?ORGID=" + encodeURIComponent(orgId) + "&sWho=" + encodeURIComponent(value));
		} else if (value === "SUB") {
			openModal("SubContractPop.asp?sUnit=" + encodeURIComponent(orgId), "", "dialogHeight:200px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", function (outValue) {
				if (field("hSubCon")) {
					field("hSubCon").value = outValue || "";
				}
			});
		}
		return false;
	};

	window.popAC = function () {
		var select = field("selAccHead");
		var issuedFor = field("hIssueToCode") ? field("hIssueToCode").value : field("hUsage") ? field("hUsage").value : "";
		var xhr;
		var root;
		if (!select) {
			return false;
		}
		select.length = Math.min(select.length, 2);
		xhr = syncGet("XMLSelectAccountHead.asp?sIssuedFor=" + encodeURIComponent(issuedFor));
		if (xhr.responseText) {
			loadXml("OutData2", xhr.responseText);
			root = xmlRoot(xmlIsland("OutData2"));
			elementChildren(root).forEach(function (node) {
				select.options[select.options.length] = new Option(attrAny(node, ["CONSUM", "Consumption", "Description"]) || getAttr(node, "AccountDescription"), attrAny(node, ["ACCHEAD", "AccountHead"]));
			});
		}
		return false;
	};

	window.CreateNew = function (obj) {
		if (obj && String(obj.selectedIndex) === "1" && field("idConsumption")) {
			field("idConsumption").style.display = "block";
		}
		return false;
	};

	window.hideDiv = function () {
		if (field("idConsumption")) {
			field("idConsumption").style.display = "none";
		}
		if (field("selAccHead")) {
			field("selAccHead").selectedIndex = 0;
		}
		return false;
	};

	window.CheckEntry = function () {
		var desc = field("txtCHead");
		var account = field("selAcc");
		var root = issueDetailsRoot();
		var doc = issueDetailsDoc();
		var node;
		var row;
		var cell;
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
		setAttr(node, "ISSFOR", field("hIssueToCode") ? field("hIssueToCode").value : field("hUsage") ? field("hUsage").value : "");
		setAttr(node, "SRC", "N");
		root.appendChild(node);
		if (document.getElementById("tblData")) {
			row = document.getElementById("tblData").insertRow(-1);
			cell = addCell(row, "ExcelDisplayCell", trim(desc.value));
			cell.align = "left";
			addCell(row, "ExcelDisplayCell", selectedText("selAcc"));
		}
		account.selectedIndex = 0;
		desc.value = "";
		return false;
	};

	window.PopDone = function () {
		var root = issueDetailsRoot();
		var select = field("selAccHead");
		syncPost("ConsumptionHeadInsert.asp", serializeXml(xmlIsland("OutData2")));
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

	window.init = function () {
		var root = xmlRoot(xmlIsland("RefData"));
		var refNode = elementChildren(root, "Ref")[0];
		var partyNode = elementChildren(root, "Party")[0];
		var issueName = getAttr(refNode, "IssName");
		if (!root) {
			return false;
		}
		if (field("hUsage")) {
			field("hUsage").value = getAttr(root, "Usage");
		}
		setDisplayText("UsageName", getAttr(root, "UsageName"));
		if (field("hIssForType")) {
			field("hIssForType").value = issueName;
		}
		if (field("hIssueToCode")) {
			field("hIssueToCode").value = getAttr(refNode, "Issue");
		}
		if (trim(issueName) && trim(issueName).toLowerCase() !== "select" && trim(issueName).toLowerCase() !== "party") {
			setDisplayText("selIssueFor", issueName);
		} else if (trim(issueName).toLowerCase() === "party") {
			setDisplayText("selIssueFor", getAttr(partyNode, "Name"));
		} else {
			setDisplayText("selIssueFor", "NA");
		}
		if (field("hPartyCode")) {
			field("hPartyCode").value = getAttr(partyNode, "Code");
		}
		return false;
	};

	window.EditUsageInfo = function () {
		var refData = xmlIsland("RefData");
		var target = /directissueitementry/i.test(window.location.pathname) ? "DirectIssueItemEntry.asp" : "MATERIALISSUEENTRY.asp";
		openModal("IssueUsageSelPop.asp?OrgID=" + encodeURIComponent(field("hUnit") ? field("hUnit").value : ""), refData, "dialogHeight:340px;dialogWidth:500px;center:yes;help:no;resizable:no;status:no", function (outValue) {
			if (getAttr(xmlRoot(outValue), "Done") === "Y") {
				syncPost("XMLSave.asp?Name=UsageSelection&SessionFlag=true", serializeXml(refData));
				form().action = target;
				form().submit();
			}
		});
		return false;
	};

	window.SelectProcess = function () {
		var orgCode = field("hUnit") ? field("hUnit").value : "";
		openModal("IssSubConProcessSelPop.asp?Unit=" + encodeURIComponent(orgCode), xmlIsland("OutData1"), "dialogHeight:370px;dialogWidth:500px;center:Yes;help:No;resizable:No;status:No", function () {
			var root = xmlRoot(xmlIsland("OutData1"));
			var node = elementChildren(root, "SubContract")[0];
			var parts = [];
			if (!node) {
				return;
			}
			parts.push(getAttr(node, "ProcessName"));
			if (getAttr(node, "LabourCharge")) {
				parts.push("LCrg.:" + getAttr(node, "LabourCharge"));
			}
			if (getAttr(node, "HardWaste")) {
				parts.push("HW%:" + getAttr(node, "HardWaste"));
			}
			if (getAttr(node, "InvWaste")) {
				parts.push("IW%:" + getAttr(node, "InvWaste"));
			}
			setDisplayText("spnProcess", parts.join(","));
		});
		return false;
	};

	window.Init = function () {
		var issueMode = field("hIssMode") ? field("hIssMode").value : "";
		var issueEntryNo = field("hIssEntryNo") ? field("hIssEntryNo").value : "";
		var appRefType = field("hAppRefType") ? field("hAppRefType").value : "";
		var appRefNo = field("hAppRefNo") ? field("hAppRefNo").value : "";
		var xhr;
		var root;
		if (trim(issueMode) === "E" && issueEntryNo) {
			xhr = syncGet("XMLGetIssueDetails.asp?IssEntNo=" + encodeURIComponent(issueEntryNo));
			if (xhr.responseText) {
				loadXml("OutData2", xhr.responseText);
				root = xmlRoot(xmlIsland("OutData2"));
				appRefType = getAttr(root, "AppRefType") || appRefType;
				appRefNo = getAttr(root, "AppRefNo") || appRefNo;
			}
		}
		if (appRefType) {
			selectOptionByValue("selRefName", appRefType);
		}
		if (appRefType && appRefNo && trim(issueMode) !== "E") {
			xhr = syncGet("InvGetItemDetForRefType.asp?RefType=" + encodeURIComponent(appRefType) + "&RefCodes=" + encodeURIComponent(appRefNo) + "&orgID=" + encodeURIComponent(field("hUnit") ? field("hUnit").value : ""));
			if (xhr.responseText) {
				loadXml("ItemData", xhr.responseText);
				elementChildren(xmlRoot(xmlIsland("ItemData"))).forEach(function (node) {
					if (/^item$/i.test(node.nodeName)) {
						appendItemDetail(node, { detectItemType: true });
					} else if (/^subcontract$/i.test(node.nodeName)) {
						issueDetailsRoot().appendChild(importForDocument(issueDetailsDoc(), node));
					}
				});
			}
		}
		window.DisplayTable(new Date());
		saveIssueDetails();
		return false;
	};

	window.PRDITEMPOPULATE = function (todaysDate) {
		return window.GetItems(todaysDate);
	};

	window.checkNumbers = checkNumbers;

	window.SetDate = function (dateValue) {
		var min = field("hMinDate") && field("hMinDate").value;
		var max = field("hMaxDate") && field("hMaxDate").value;
		if (dateDiffDays(min, dateValue) < 0 || dateDiffDays(dateValue, max) < 0) {
			setDatePickerValue(max);
		} else {
			setDatePickerValue(dateValue);
		}
	};

	window.MinDate = function () {
		var min = field("hMinDate") && field("hMinDate").value;
		var max = field("hMaxDate") && field("hMaxDate").value;
		var value = getDatePickerValue();
		if (dateDiffDays(min, value) < 0 || dateDiffDays(value, max) < 0) {
			alert("Issue Date Should Be With in the Financial Year " + min + " and " + max);
			setDatePickerValue(max);
		}
	};

	window.Back = function () {
		form().action = "MRSMGMTLIST.ASP?hCheck=I";
		form().submit();
	};

	window.DisplayDet = function () {};
	window.LoadData = function () {};

	window.DisplayStock = function (obj) {
		var parts = String(obj && obj.name || "").split("Z");
		var tempValues = (parts[2] || "") + ":" + (parts[1] || "") + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (parts[3] || "") + ":" + (parts[4] || "") + ":" + (parts[5] || "") + ":" + (parts[6] || "");
		openModal("itmStockPoP.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:450px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No");
	};

	window.GetAddDetails = function (orgId, classCode, itemCode, mrsNo) {
		var qty = field("txtQtyPPX" + classCode + "x" + itemCode) || field("txtQtyPPX" + classCode + "X" + itemCode);
		var tempValues = classCode + "|" + itemCode + "|" + orgId + "|" + mrsNo + "|" + (qty ? qty.value : "");
		openModal("IssueAddDetails.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:400px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No");
	};

	window.CheckSch = function (obj) {
		var parts = String(obj && obj.name || "").split(":");
		var itemCode = parts[2] || "";
		var classCode = parts[1] || "";
		var entryNo = parts[3] || "";
		var optionName = parts[4] || "";
		var qty = field("txtQtyA" + classCode + "A" + itemCode + "A" + entryNo);
		var tempValues = (qty ? qty.value : "") + ":" + itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (field("hOrgID") ? field("hOrgID").value : "") + ":" + entryNo + ":" + optionName;
		openModal("mrsIssueSchedulePoP.asp?sTemp=" + encodeURIComponent(tempValues), xmlIsland("Data"), "dialogHeight:480px;dialogWidth:390px;center:Yes;help:No;resizable:No;status:No");
	};

	window.GetSch = function (obj) {
		var parts = String(obj && obj.name || "").split("X");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var sch = field("hSchA" + classCode + "A" + itemCode + "A" + entryNo);
		var tempValues;
		if (!sch || sch.value !== "S" || trim(obj.value) === "" || toNumber(obj.value) === 0) {
			return;
		}
		tempValues = obj.value + ":" + itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (field("hOrgID") ? field("hOrgID").value : "") + ":" + entryNo;
		openModal("mrsIssueScheduleEntryPoP.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:490px;dialogWidth:330px;center:Yes;help:No;resizable:No;status:No");
	};

	window.CheckQty = function (obj) {
		var parts = String(obj && obj.name || "").split(":");
		var tempValues = (parts[2] || "") + ":" + (parts[1] || "") + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (parts[3] || "") + ":" + (parts[4] || "");
		openModal("mrsIssueQtyParaPoP.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:385px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No");
	};

	window.DisplayItem = function (obj) {
		var parts = String(obj || "").split("A");
		var tempValues = parts.length > 3 ? parts[0] + "A" + parts[2] + "A" + parts[1] + "A" + parts[3] : obj;
		openModal("itmDetailsPop.asp?sTemp=" + encodeURIComponent(tempValues), "", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No");
	};

	window.CheckLot = function (obj, issueQty, attributeList) {
		var parts = String(obj && obj.name || "").split("Z");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var optionName = parts[4] || "";
		var hasStoreBin = parts.length > 7;
		var locCode = hasStoreBin ? parts[5] || "" : entryNo;
		var binCode = hasStoreBin ? parts[6] || "" : optionName;
		var orgId = hasStoreBin ? field("hOrgID") && field("hOrgID").value || parts[5] || "" : parts[5] || "";
		var parsedAttrs = parseAttributeList(attributeList || (hasStoreBin ? parts.slice(7).join("Z") : parts[6]) || "");
		var issueType = issueTypeValue();
		var packFlag = field("hPickPackFlag") ? field("hPickPackFlag").value : "";
		var usage = field("hUsage") ? field("hUsage").value : "";
		var popupName = islandExists("OutData2") ? "mrsPickDetailPoP.asp" : "MRPickDetPoP.asp";
		var tempValues = itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + entryNo + ":" + optionName + ":" + locCode + ":" + binCode + ":" + (issueQty || "") + ":" + issueType + ":" + usage + ":" + orgId + ":" + parsedAttrs.attributeList + ":" + parsedAttrs.attributeId;
		openModal(popupName + "?sTemp=" + encodeURIComponent(tempValues) + "&AttributeList=" + encodeURIComponent(attributeList || "") + "&PickPackFlag=" + encodeURIComponent(packFlag), xmlIsland("OutData1"), "dialogHeight:350px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No", function () {
			updatePickQuantity(itemCode, classCode, entryNo);
		});
		return false;
	};

	window.CheckST = function (obj) {
		var parts = String(obj && obj.name || "").split(":");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var optionName = parts[4] || "";
		var orgId = parts[5] || "";
		var store = field("hStoName" + classCode + "A" + itemCode + "A" + entryNo);
		var tempValues = itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (store ? store.value : "") + ":" + entryNo + ":" + optionName + ":" + orgId;
		openModal("mrsIssueSTPoP.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:330px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function () {
			var root = detailsRoot();
			var item = matchingItem(root, entryNo, itemCode, classCode);
			var qty = field("txtQtyTraX" + classCode + "X" + itemCode + "X" + entryNo) || field("txtQtyTraX" + classCode + "X" + itemCode);
			if (item && qty) {
				qty.value = getAttr(item, "TRAQTY") || getAttr(item, "ISSQTY");
			}
		});
	};

	window.CheckSTPRSch = function (obj, todaysDate, target) {
		var parts = String(obj && obj.name || "").split("Z");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var optionName = parts[4] || "";
		var qty = target === "ST" ? field("txtQtyTraX" + classCode + "X" + itemCode + "X" + entryNo) : field("txtQtyPrX" + classCode + "X" + itemCode + "X" + entryNo);
		var root = detailsRoot();
		var item = matchingItem(root, entryNo, itemCode, classCode);
		var doc = xmlDocument(detailsIsland());
		var scheduleName = target === "ST" ? "STSchedule" : "PRSchedule";
		var value;
		var node;
		if (!qty || trim(qty.value) === "") {
			alert("Enter Quantity");
			if (qty) {
				qty.focus();
			}
			obj.selectedIndex = 0;
			return false;
		}
		if (!checkNumbers(qty.value)) {
			alert("Enter Numerals Only");
			qty.focus();
			obj.selectedIndex = 0;
			return false;
		}
		if (toNumber(qty.value) === 0) {
			obj.selectedIndex = 0;
			return false;
		}
		if (!item && root && doc) {
			item = createElement(doc, "ITEM");
			setAttr(item, "ENTRYNO", entryNo);
			setAttr(item, "ITEMCODE", itemCode);
			setAttr(item, "ITMCODE", itemCode);
			setAttr(item, "CLASSCODE", classCode);
			setAttr(item, "CLACODE", classCode);
			root.appendChild(item);
		}
		if (!item) {
			return false;
		}
		clearChildrenByName(item, "STSchedule");
		clearChildrenByName(item, "PRSchedule");
		clearChildrenByName(item, "ScheduleDetails");
		if (obj.selectedIndex === 1 || obj.selectedIndex === 2) {
			value = prompt(obj.selectedIndex === 1 ? "Enter No of Days" : "Enter the Date", obj.selectedIndex === 1 ? "0" : "");
			if (value == null || trim(value) === "") {
				obj.selectedIndex = 0;
				return false;
			}
			if (obj.selectedIndex === 1 && !checkNumbers(value)) {
				alert("Enter Numerals Only");
				obj.selectedIndex = 0;
				return false;
			}
			if (obj.selectedIndex === 2 && (!parseDate(value) || dateDiffDays(todaysDate, value) < 0)) {
				alert("Invalid Date");
				obj.selectedIndex = 0;
				return false;
			}
			node = createElement(doc, scheduleName);
			setAttr(node, "STYPE", selectedValue(obj.name));
			setAttr(node, "SVALUE", value);
			item.appendChild(node);
		} else if (obj.selectedIndex === 3) {
			node = createElement(doc, scheduleName);
			setAttr(node, "STYPE", selectedValue(obj.name));
			setAttr(node, "SVALUE", "");
			item.appendChild(node);
			openModal((target === "ST" ? "mrsSTSchedulePoP.asp" : "mrsPRSchedulePoP.asp") + "?sTemp=" + encodeURIComponent(qty.value + ":" + itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + entryNo + ":" + optionName), detailsIsland(), "dialogHeight:510px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No");
		}
		return false;
	};

	window.RemoveXML = function () {
		var root = detailsRoot();
		var items = elementChildren(root, "ITEMDETAILS");
		for (var i = 0; i < items.length; i += 1) {
			while (items[i].firstChild) {
				items[i].removeChild(items[i].firstChild);
			}
		}
	};

	window.CheckSubmit = function (todaysDate) {
		var root = buildIssueXml(todaysDate);
		var issueMode = field("hIssMode") ? field("hIssMode").value : "N";
		var issueEntryNo = field("hIssEntryNo") ? field("hIssEntryNo").value : "";
		var callFrom = field("hCallFrom") ? field("hCallFrom").value : "";
		if (!root) {
			return false;
		}
		decorateIssueRoot(root);
		syncPost(issueMode === "E" ? "XMLSave.asp?Name=mrsIssueDataEdit&SessionFlag=true" : "XMLSave.asp?Name=mrsIssueData&SessionFlag=true", serializeXml(xmlIsland("OutData1")));
		form().action = issueMode === "E" ? "mrsIssueUpdate.asp?hCallFrom=" + encodeURIComponent(callFrom) + "&IssEntNo=" + encodeURIComponent(issueEntryNo) : "mrsIssueInsert.asp?hCallFrom=" + encodeURIComponent(callFrom);
		form().submit();
		return false;
	};
}(window, document));
