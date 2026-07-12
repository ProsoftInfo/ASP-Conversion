(function (window, document) {
	"use strict";

	var buttonPressed = "";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.FormName || document.forms.formname || document.forms.FormName || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function fieldValue(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setFieldValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
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

	function dialogArgs() {
		var args = window.dialogArguments;
		var match;
		var id;
		if (!args) {
			match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
			id = match ? decodeURIComponent(match[1]) : "";
			if (id && window.opener && window.opener.__itmsDialogArgs) {
				args = window.opener.__itmsDialogArgs[id];
				window.dialogArguments = args;
			}
		}
		return args || null;
	}

	function xmlDocument(object) {
		return object && object.XMLDocument || object && object._doc || object || null;
	}

	function dialogDocument() {
		return xmlDocument(dialogArgs());
	}

	function root() {
		var args = dialogArgs();
		var doc = xmlDocument(args);
		return args && args.documentElement || doc && doc.documentElement || null;
	}

	function createNode(name) {
		var doc = dialogDocument();
		if (doc && doc.createElement) {
			return doc.createElement(name);
		}
		return document.implementation.createDocument("", "", null).createElement(name);
	}

	function xmlIsland(id) {
		var island = byId(id);
		return xmlDocument(island);
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

	function returnValue(value) {
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

	function requestText(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		if (xhr.status >= 400) {
			throw new Error("Request failed: " + url + " (" + xhr.status + ")");
		}
		return xhr.responseText || "";
	}

	function loadXmlText(doc, xmlText) {
		if (!doc || !trim(xmlText)) {
			return false;
		}
		if (doc.loadXML) {
			doc.loadXML(xmlText);
			return true;
		}
		return false;
	}

	function loadXmlFromUrl(url, islandId, alertOnText) {
		var text = requestText(url);
		var doc = xmlIsland(islandId);
		if (trim(text) && loadXmlText(doc, text)) {
			return doc;
		}
		if (alertOnText && trim(text)) {
			alert(text);
		}
		return doc;
	}

	function selectMode() {
		return trim(fieldValue("hSelectMode")).toUpperCase();
	}

	function eventKey(eventArg) {
		var evt = eventArg || {};
		return evt.keyCode || evt.which || 0;
	}

	function encodeQuery(value) {
		return encodeURIComponent(String(value == null ? "" : value));
	}

	function queryParam(name) {
		var pattern = new RegExp("[?&]" + name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "=([^&]*)", "i");
		var match = String(window.location.search || "").match(pattern);
		return match ? decodeURIComponent(match[1].replace(/\+/g, " ")) : "";
	}

	function hiddenValue(prefix, itemCode, classCode) {
		return fieldValue(prefix + itemCode + classCode);
	}

	function valueOrZero(value) {
		return trim(value) || "0";
	}

	function selectedItems() {
		return childElements(root(), "Item");
	}

	function selectedMaterials() {
		var result = [];
		childElements(root(), "Materials").forEach(function (materials) {
			result = result.concat(childElements(materials, "Entry"));
		});
		return result;
	}

	function nextEntryNo() {
		var max = 0;
		selectedItems().forEach(function (item) {
			var value = parseInt(attr(item, "EntryNo"), 10);
			if (!isNaN(value) && value > max) {
				max = value;
			}
		});
		return max + 1;
	}

	function removeAllItems() {
		var currentRoot = root();
		if (!currentRoot) {
			return;
		}
		selectedItems().forEach(function (item) {
			currentRoot.removeChild(item);
		});
	}

	function itemExists(parts, attributeList, checkAttributeList) {
		var itemCode = trim(parts[5]);
		var classCode = trim(parts[6]);
		return selectedItems().some(function (item) {
			var sameItem = trim(attr(item, "ItemCode")) === itemCode && trim(attr(item, "ClassCode")) === classCode;
			return checkAttributeList ? sameItem && trim(attr(item, "AttributeList")) === trim(attributeList) : sameItem;
		});
	}

	function removeItem(parts) {
		var currentRoot = root();
		var itemCode = trim(parts[5] != null ? parts[5] : parts[1]);
		var classCode = trim(parts[6] != null ? parts[6] : parts[2]);
		selectedItems().forEach(function (item) {
			if (trim(attr(item, "ItemCode")) === itemCode && trim(attr(item, "ClassCode")) === classCode) {
				currentRoot.removeChild(item);
			}
		});
	}

	function decodeItemName(value) {
		return String(value == null ? "" : value).replace(/-/g, " ");
	}

	function getBomItem(itemCode, itemNode) {
		var bomDoc;
		try {
			bomDoc = loadXmlFromUrl("GetBoMItemDet.asp?ItemCode=" + encodeQuery(itemCode), "BOMItem", false);
			if (bomDoc && bomDoc.documentElement && itemNode) {
				itemNode.appendChild(bomDoc.documentElement.cloneNode(true));
			}
		} catch (error) {
			alert(error.message || error);
		}
	}

	function appendItem(parts, attributeList, attributeText, includeSupplierValues) {
		var currentRoot = root();
		var node;
		var itemCode = parts[5] || "";
		var classCode = parts[6] || "";
		if (!currentRoot) {
			return null;
		}
		node = createNode("Item");
		setAttr(node, "EntryNo", nextEntryNo());
		setAttr(node, "CompanyItemCode", parts[0] || "");
		setAttr(node, "ItemCode", itemCode);
		setAttr(node, "ClassCode", classCode);
		setAttr(node, "ItemName", decodeItemName(parts[1]) + (attributeText || ""));
		setAttr(node, "ClassName", parts[2] || "");
		setAttr(node, "StoresUoM", parts[4] || "");
		setAttr(node, "Decimal", parts[7] || "");
		setAttr(node, "ReceiptNum", parts[8] || "");
		setAttr(node, "AttributeList", attributeList || "");
		setAttr(node, "ItemRate", valueOrZero(hiddenValue("hItemRate", itemCode, classCode)));
		setAttr(node, "ItemStock", valueOrZero(hiddenValue("hItemStock", itemCode, classCode)));
		setAttr(node, "LocAndBinCount", valueOrZero(hiddenValue("hBinAndLocCheck", itemCode, classCode)));
		setAttr(node, "LocNo", valueOrZero(hiddenValue("hLocNo", itemCode, classCode)));
		setAttr(node, "BinNo", valueOrZero(hiddenValue("hBinNo", itemCode, classCode)));
		setAttr(node, "PartyCode", fieldValue("hPartyCode"));
		setAttr(node, "PartyType", "");
		setAttr(node, "PartySubType", "");
		setAttr(node, "SuppItemCode", includeSupplierValues ? hiddenValue("hSuppItemCode", itemCode, classCode) : "");
		setAttr(node, "SuppItemDesc", includeSupplierValues ? hiddenValue("hSuppItemDesc", itemCode, classCode) : "");
		setAttr(node, "MarketPrice", valueOrZero(hiddenValue("hMarketPrice", itemCode, classCode)));
		currentRoot.appendChild(node);
		getBomItem(itemCode, node);
		return node;
	}

	function attributeDocForItem(itemCode) {
		try {
			return loadXmlFromUrl("XMLGetAttributeListForItem.asp?ItemCode=" + encodeQuery(itemCode), "XMLAttributeList", false);
		} catch (error) {
			alert(error.message || error);
			return null;
		}
	}

	function collectAttributeList(parts, checkStock) {
		var itemCode = parts[5] || "";
		var classCode = parts[6] || "";
		var hasAttributes = fieldValue("hAttribValZ" + itemCode + "Z" + classCode);
		var doc;
		var rootNode;
		var values = [];
		var names = [];
		var stockValues = [];
		if (!trim(hasAttributes) || trim(hasAttributes).toUpperCase() === "N") {
			return { ok: true, value: "", text: "" };
		}
		doc = attributeDocForItem(itemCode);
		rootNode = doc && doc.documentElement;
		childElements(rootNode, "Attribute").forEach(function (attribute) {
			var id = attr(attribute, "ID");
			var select = field("SelAttributeZ" + itemCode + "Z" + classCode + "Z" + id);
			var selected = select ? select.options[select.selectedIndex] : null;
			var selectedValue = selected ? selected.value : "";
			var valueParts;
			var idParts;
			values.push(selectedValue);
			valueParts = String(selectedValue || "").split(":");
			idParts = String(valueParts[0] || "").split("#");
			if (trim(idParts[1]) !== "0") {
				if (trim(valueParts[1])) {
					names.push(trim(valueParts[1]));
				}
				if (trim(idParts[1])) {
					stockValues.push(trim(idParts[1]));
				}
			}
		});
		if (checkStock && stockValues.length) {
			if (!checkAttributeStock(itemCode, stockValues.join(","))) {
				return { ok: false, value: "", text: "" };
			}
		}
		return {
			ok: true,
			value: values.join(",").replace(/'/g, "~~"),
			text: names.length ? "[" + names.join(",") + "]" : ""
		};
	}

	function checkAttributeStock(itemCode, stockAttributeList) {
		var callFrom = trim(fieldValue("hCallFrom")).toUpperCase();
		var text;
		try {
			text = trim(requestText("GetItemAttributeStock.asp?ItemCode=" + encodeQuery(itemCode) + "&AttID=" + encodeQuery(stockAttributeList)));
		} catch (error) {
			alert(error.message || error);
			return false;
		}
		if (!text) {
			return true;
		}
		if (text.length === 1) {
			if (text === "N" && callFrom !== "PUR" && callFrom !== "SO" && callFrom !== "MR") {
				alert("The selected attribute has no stock");
				return false;
			}
			return true;
		}
		alert(text);
		return false;
	}

	function addXmlFun(obj) {
		var parts = String(obj && obj.value || "").split(":");
		var itemCode = parts[5] || "";
		var classCode = parts[6] || "";
		var hasAttributes = trim(fieldValue("hAttribValZ" + itemCode + "Z" + classCode));
		var addButton = field("btnAddToList");
		if (!hasAttributes || hasAttributes.toUpperCase() === "N") {
			if (addButton) {
				addButton.disabled = true;
			}
			xmlFun(obj);
			return;
		}
		if (addButton) {
			addButton.disabled = false;
		}
	}

	function addFun() {
		var count = parseInt(fieldValue("hChkCount"), 10) || 0;
		var check;
		var parts;
		var attributes;
		var i;
		for (i = 1; i <= count; i += 1) {
			check = field("ChkZ" + i);
			if (check && check.checked) {
				check.checked = false;
				parts = String(check.value || "").split(":");
				attributes = collectAttributeList(parts, true);
				if (!attributes.ok) {
					return;
				}
				if (selectMode() === "M") {
					if (!itemExists(parts, attributes.value, true)) {
						appendItem(parts, attributes.value, attributes.text, true);
					}
				} else {
					removeAllItems();
					appendItem(parts, attributes.value, attributes.text, false);
				}
			}
		}
		if (field("btnAddToList")) {
			field("btnAddToList").disabled = true;
		}
		displayList();
	}

	function xmlFun(obj) {
		var parts = String(obj && obj.value || "").split(":");
		var attributes = collectAttributeList(parts, false);
		if (!attributes.ok) {
			return;
		}
		if (selectMode() === "M") {
			if (obj && obj.checked) {
				if (!itemExists(parts, attributes.value, false)) {
					appendItem(parts, attributes.value, attributes.text, true);
				}
			} else {
				removeItem(parts);
			}
			displayList();
			return;
		}
		removeAllItems();
		if (obj && obj.checked) {
			appendItem(parts, attributes.value, attributes.text, false);
		}
		displayList();
	}

	function escapeHtml(value) {
		return String(value == null ? "" : value)
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;")
			.replace(/'/g, "&#39;");
	}

	function displayList() {
		var html = '<br><TABLE class="TableOutLineOnly" cellspacing="1" width="100%">';
		selectedItems().forEach(function (item) {
			var value = [
				attr(item, "CompanyItemCode"),
				attr(item, "ItemCode"),
				attr(item, "ClassCode"),
				attr(item, "ClassName"),
				attr(item, "ItemName").replace(/~~/g, '"')
			].join(":");
			html += '<tr><td class="ExcelDisplayCell">';
			html += '<input type="checkbox" name="chk" value="' + escapeHtml(value) + '" checked onclick="RemoveNode(this)">';
			html += '</td><td class="ExcelDisplayCell">' + escapeHtml(attr(item, "CompanyItemCode")) + '</td>';
			html += '<td class="ExcelDisplayCell">' + escapeHtml(attr(item, "ItemName").replace(/~~/g, '"')) + '</td>';
			html += '<td class="ExcelDisplayCell">' + escapeHtml(attr(item, "ClassName")) + '</td></tr>';
		});
		selectedMaterials().forEach(function (entry) {
			html += '<tr><td class="ExcelDisplayCell">';
			html += '<input type="checkbox" name="chk" value="' + escapeHtml(attr(entry, "SlNo")) + '" checked onclick="RemoveNode(this)">';
			html += '</td><td class="ExcelDisplayCell">--NA--</td>';
			html += '<td class="ExcelDisplayCell">' + escapeHtml(attr(entry, "ItemName")) + '</td>';
			html += '<td class="ExcelDisplayCell">--NA--</td></tr>';
		});
		html += "</table><br>";
		if (byId("idSelList")) {
			byId("idSelList").innerHTML = html;
		}
	}

	function sendValue() {
		buttonPressed = "Done";
		setAttr(root(), "Action", "Done");
		returnValue(root());
		window.close();
	}

	function setCell(cell, text, align) {
		cell.className = "ExcelDisplayCell";
		cell.align = align || "Left";
		cell.innerHTML = escapeHtml(text);
	}

	function setSelectCell(cell, item, tabIndex) {
		var input = document.createElement("input");
		var itemValue = [
			item.companyItemCode,
			String(item.itemName || "").replace(/ /g, "-").replace(/"/g, "~~").replace(/'/g, "~"),
			String(item.className || "").replace(/ /g, "-"),
			item.stock,
			item.uom,
			item.itemCode,
			item.classCode,
			item.decimalAllowed,
			item.receiptNum,
			item.attributeList
		].join(":");
		input.type = selectMode() === "M" ? "checkbox" : "radio";
		input.name = "ChkZ" + (selectMode() === "M" ? item.sNo : item.counter);
		input.value = itemValue;
		input.onclick = selectMode() === "M" ? function () { addXmlFun(input); } : function () { xmlFun(input); };
		if (tabIndex) {
			input.tabIndex = tabIndex;
		}
		cell.className = "ExcelDisplayCell";
		cell.align = "center";
		cell.appendChild(input);
	}

	function addHidden(cell, name, value) {
		var input = document.createElement("input");
		input.type = "hidden";
		input.name = name;
		input.value = value == null ? "" : String(value);
		cell.appendChild(input);
	}

	function setUomCell(cell, item) {
		cell.className = "ExcelDisplayCell";
		cell.align = "Center";
		cell.appendChild(document.createTextNode(item.uom || ""));
		addHidden(cell, "hItemRate" + item.itemCode + item.classCode, item.itemRate);
		addHidden(cell, "hItemStock" + item.itemCode + item.classCode, item.stock);
		addHidden(cell, "hBinAndLocCheck" + item.itemCode + item.classCode, item.locBinCount);
		addHidden(cell, "hLocNo" + item.itemCode + item.classCode, item.locNo);
		addHidden(cell, "hBinNo" + item.itemCode + item.classCode, item.binNo);
		addHidden(cell, "hMarketPrice" + item.itemCode + item.classCode, item.marketPrice);
		addHidden(cell, "hSuppItemCode" + item.itemCode + item.classCode, item.partyItemCode);
		addHidden(cell, "hSuppItemDesc" + item.itemCode + item.classCode, item.partyItemDesc);
	}

	function setAttributeCell(cell, item) {
		var doc = attributeDocForItem(item.itemCode);
		var rootNode = doc && doc.documentElement;
		var attributes = childElements(rootNode, "Attribute");
		cell.className = "ExcelDisplayCell";
		if (!attributes.length) {
			cell.innerHTML = "N/A";
			addHidden(cell, "hAttribValZ" + item.itemCode + "Z" + item.classCode, "N");
			return;
		}
		attributes.forEach(function (attribute) {
			var id = attr(attribute, "ID");
			var select = document.createElement("select");
			var defaultOption = document.createElement("option");
			select.name = "SelAttributeZ" + item.itemCode + "Z" + item.classCode + "Z" + id;
			select.className = "FormElem";
			defaultOption.value = id + "#0:" + attr(attribute, "Name");
			defaultOption.text = attr(attribute, "Name");
			select.add(defaultOption);
			childElements(attribute).forEach(function (optionNode) {
				var option = document.createElement("option");
				option.value = id + "#" + attr(optionNode, "Value") + ":" + attr(optionNode, "Name");
				option.text = attr(optionNode, "Name");
				select.add(option);
			});
			cell.appendChild(select);
		});
		addHidden(cell, "hAttribValZ" + item.itemCode + "Z" + item.classCode, "Y");
	}

	function itemFromNode(node) {
		return {
			sNo: attr(node, "SNO"),
			counter: attr(node, "Counter"),
			companyItemCode: trim(attr(node, "ComItemCode")).replace(/ /g, ""),
			itemName: attr(node, "ItemName"),
			partyItemCode: attr(node, "PartyItemCode"),
			partyItemDesc: attr(node, "PartyItemDesc"),
			className: attr(node, "ClassName"),
			stock: attr(node, "Stock"),
			uom: attr(node, "UOM"),
			itemCode: attr(node, "ItemCode"),
			classCode: attr(node, "ClassCode"),
			decimalAllowed: attr(node, "DecimalAllowed"),
			receiptNum: attr(node, "ReceiptNum"),
			attributeList: attr(node, "AttributeList"),
			itemRate: attr(node, "ItemRate"),
			locBinCount: attr(node, "hBinAndLocCheck"),
			locNo: attr(node, "hLocNo"),
			binNo: attr(node, "hBinNo"),
			marketPrice: attr(node, "hMarketPrice")
		};
	}

	function clearTable() {
		var table = byId("tblItem");
		if (!table) {
			return;
		}
		while (table.rows.length > 5) {
			table.deleteRow(2);
		}
	}

	function insertDataRow() {
		var table = byId("tblItem");
		return table.insertRow(Math.max(2, table.rows.length - 3));
	}

	function renderNoRecords(tabIndex) {
		var row = insertDataRow();
		var cell = row.insertCell();
		cell.innerText = "No Records Found";
		cell.className = "ExcelDisplayCell";
		cell.align = "center";
		cell.colSpan = 7;
		setFieldValue("txtCurrPage", "0");
		setFieldValue("hPage", "0");
		if (byId("spanTotPage")) {
			byId("spanTotPage").innerText = "0";
		}
		setFooterTabIndexes(tabIndex, "0");
	}

	function setFooterTabIndexes(tabIndex, currentPage) {
		if (field("txtCurrPage")) {
			field("txtCurrPage").tabIndex = tabIndex;
			setFieldValue("txtCurrPage", currentPage);
		}
		if (fieldValue("hButtDispMode") === "Y" && field("btnAddNew")) {
			tabIndex += 1;
			field("btnAddNew").tabIndex = tabIndex;
		}
		if (field("btnAddToList")) {
			tabIndex += 1;
			field("btnAddToList").tabIndex = tabIndex;
		}
		if (field("btnDone")) {
			tabIndex += 1;
			field("btnDone").tabIndex = tabIndex;
		}
	}

	function renderTempItems(updateFooterTabs) {
		var doc = xmlIsland("TempItem");
		var tempRoot = doc && doc.documentElement;
		var items = childElements(tempRoot, "Item");
		var currentPage = attr(tempRoot, "CurrPage") || "0";
		var totalPage = attr(tempRoot, "TotPage") || "0";
		var lastSlNo = "0";
		var tabIndex = 6;
		clearTable();
		if (!items.length) {
			renderNoRecords(tabIndex);
			return;
		}
		items.forEach(function (node) {
			var item = itemFromNode(node);
			var row = insertDataRow();
			lastSlNo = item.sNo || lastSlNo;
			setSelectCell(row.insertCell(), item, updateFooterTabs ? tabIndex : 0);
			setCell(row.insertCell(), item.companyItemCode, "Left");
			setCell(row.insertCell(), item.itemName, "Left");
			setCell(row.insertCell(), item.partyItemCode, "Left");
			setCell(row.insertCell(), item.partyItemDesc, "Left");
			setCell(row.insertCell(), item.stock, "Center");
			setCell(row.insertCell(), item.className, "Center");
			setUomCell(row.insertCell(), item);
			setAttributeCell(row.insertCell(), item);
			tabIndex += 1;
		});
		setFieldValue("hChkCount", lastSlNo);
		setFieldValue("txtCurrPage", currentPage);
		setFieldValue("hPage", currentPage);
		if (byId("spanTotPage")) {
			byId("spanTotPage").innerText = totalPage;
		}
		if (updateFooterTabs) {
			setFooterTabIndexes(tabIndex, currentPage);
		}
		displayList();
	}

	function searchQuery(args) {
		var check = field("chkExact") && field("chkExact").checked ? "Y" : "N";
		return String(args || "") +
			"&SICode=" + encodeQuery(fieldValue("txtSearchItemCode")) +
			"&SIName=" + encodeQuery(fieldValue("txtSearchItemName")) +
			"&SPCode=" + encodeQuery(fieldValue("txtSearchPartyItemCode")) +
			"&SPName=" + encodeQuery(fieldValue("txtSearchPartyItemName")) +
			"&CheckStart=" + encodeQuery(check) +
			"&hStock=" + encodeQuery(fieldValue("txtSearchStock"));
	}

	function showPage(args) {
		var query = searchQuery(args);
		buttonPressed = "Page";
		setAttr(root(), "Action", "Page");
		setAttr(root(), "PassQuery", query);
		try {
			loadXmlFromUrl("XMLGetItemSelectRel.asp?" + query, "TempItem", true);
			renderTempItems(true);
		} catch (error) {
			alert(error.message || error);
		}
	}

	function callChangePage(eventArg) {
		var key = eventKey(eventArg);
		var page = parseFloat(fieldValue("txtCurrPage")) || 0;
		var lastPage = parseFloat(byId("spanTotPage") ? byId("spanTotPage").innerText : "0") || 0;
		var request = fieldValue("hRequest");
		if (key === 33) {
			if (page > 1) {
				page -= 1;
			}
			setFieldValue("hPage", page);
			showPage(request + "&Page=" + page);
		} else if (key === 34) {
			if (page < lastPage) {
				page += 1;
			}
			setFieldValue("hPage", page);
			showPage(request + "&Page=" + page);
		} else if (key >= 48 && key <= 57) {
			setFieldValue("hPage", page);
			showPage(request + "&Page=" + page);
		}
	}

	function callSearchMain(eventArg) {
		var key = eventKey(eventArg);
		var request = fieldValue("hRequest");
		var page = parseFloat(fieldValue("hPage")) || 0;
		var lastPage = parseFloat(byId("spanTotPage") ? byId("spanTotPage").innerText : "0") || 0;
		if (key === 13) {
			page = fieldValue("txtCurrPage");
			showPage(request + "&Page=" + page);
		} else if (key === 33) {
			if (page > 1) {
				page -= 1;
			}
			showPage(request + "&Page=" + page);
		} else if (key === 34) {
			if (page < lastPage) {
				page += 1;
			}
			showPage(request + "&Page=" + page);
		}
	}

	function callSearch() {
		var request = fieldValue("hRequest");
		var page = fieldValue("txtCurrPage");
		showPage(request + "&Page=" + page);
	}

	function callSearchStock(eventArg) {
		var key = eventKey(eventArg);
		var request = fieldValue("hRequest");
		var page = fieldValue("txtCurrPage");
		if ((key >= 48 && key <= 57) || key === 60 || key === 61 || key === 62 || !key) {
			showPage(request + "&Page=" + page);
		}
	}

	function init() {
		buttonPressed = "";
		try {
			loadXmlFromUrl("XMLGetItemSelectRel.asp?" + fieldValue("hTemp"), "TempItem", true);
			renderTempItems(false);
		} catch (error) {
			alert(error.message || error);
		}
	}

	function removeMaterial(slNo) {
		childElements(root(), "Materials").forEach(function (materials) {
			childElements(materials, "Entry").forEach(function (entry) {
				if (trim(attr(entry, "SlNo")) === trim(slNo)) {
					materials.removeChild(entry);
				}
			});
		});
	}

	function removeNode(obj) {
		var value = String(obj && obj.value || "");
		var parts = value.split(":");
		if (obj && obj.checked) {
			return;
		}
		if (parts.length === 1) {
			removeMaterial(value);
			displayList();
			return;
		}
		selectedItems().forEach(function (item) {
			if (trim(attr(item, "ItemCode")) === trim(parts[1]) && trim(attr(item, "ClassCode")) === trim(parts[2])) {
				root().removeChild(item);
			}
		});
		asArray(form() && form().elements).forEach(function (element) {
			var checkParts;
			if (element.type === "checkbox" && /^ChkZ/.test(element.name || "")) {
				checkParts = String(element.value || "").split(":");
				if (trim(parts[1]) === trim(checkParts[5]) && trim(parts[2]) === trim(checkParts[6])) {
					element.checked = false;
				}
			}
		});
		displayList();
	}

	function withOutMat() {
		var orgId = fieldValue("hOrgID");
		var flag = queryParam("Flag") || queryParam("flag");
		var url = "../Purchase/Transaction/SelMaterialNew.asp?orgID=" + encodeQuery(orgId) + "&hSelectMode=M&Flag=" + encodeQuery(flag);
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, dialogArgs(), "dialogHeight:600px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", function (outRoot) {
				if (!outRoot) {
					return;
				}
				childElements(outRoot, "Materials").forEach(function (materials) {
					root().appendChild(materials.cloneNode(true));
				});
				displayList();
			});
		} else {
			window.open(url, "_blank", "height=600,width=550,resizable=no,status=no");
		}
	}

	function windowOnUnload() {
		returnValue(root());
	}

	window.AddXmlFun = addXmlFun;
	window.AddFun = addFun;
	window.XmlFun = xmlFun;
	window.XMLFun = xmlFun;
	window.DispList = displayList;
	window.SendValue = sendValue;
	window.showpage = showPage;
	window.ShowPage = showPage;
	window.CallChangePage = callChangePage;
	window.CallSearchMain = callSearchMain;
	window.CallSearch = callSearch;
	window.CallSearchStock = callSearchStock;
	window.Init = init;
	window.ClearTable = clearTable;
	window.RemoveNode = removeNode;
	window.WithOutMat = withOutMat;
	window.window_onunload = windowOnUnload;

	window.addEventListener("beforeunload", windowOnUnload);
}(window, document));
