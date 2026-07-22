(function (window, document) {
	"use strict";

	var rowCount = 0;

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
		var controls = frm && frm.elements || [];
		if (controls[name]) {
			return controls[name];
		}
		for (var i = 0; i < controls.length; i += 1) {
			if (String(controls[i].name || "").toLowerCase() === String(name).toLowerCase()) {
				return controls[i];
			}
		}
		return document.getElementById(name) || null;
	}

	function setText(id, value) {
		var node = document.getElementById(id) || window[id];
		if (node) {
			node.textContent = value == null ? "" : String(value);
			node.innerText = node.textContent;
			node.innerHTML = node.textContent;
		}
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(trim(value));
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

	function selectedText(name) {
		var control = field(name);
		return control && control.options && control.selectedIndex >= 0 ? control.options[control.selectedIndex].text : "";
	}

	function selectByValue(name, value) {
		var control = field(name);
		if (!control || !control.options) {
			return;
		}
		for (var i = 0; i < control.options.length; i += 1) {
			if (trim(control.options[i].value) === trim(value)) {
				control.selectedIndex = i;
				return;
			}
		}
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

	function outDoc() {
		return xmlDocument(xmlIsland("OutData2"));
	}

	function outRoot() {
		return xmlRoot(xmlIsland("OutData2"));
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

	function descendants(node, name) {
		var result = [];
		var wanted = String(name).toLowerCase();
		function walk(current) {
			elementChildren(current).forEach(function (child) {
				if (String(child.nodeName).toLowerCase() === wanted) {
					result.push(child);
				}
				walk(child);
			});
		}
		walk(node);
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

	function loadXmlResponse(islandName, xhr) {
		if (!xhr || xhr.status >= 400 || !trim(xhr.responseText)) {
			return false;
		}
		loadXml(islandName, xhr.responseText);
		return true;
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		window.open(url, "_blank");
		return null;
	}

	function popupSize(type, fallbackProgram, fallbackHeight, fallbackWidth) {
		var parts = typeof window.GetWindowSizeForPopup === "function" ? String(window.GetWindowSizeForPopup(type)).split(":") : [];
		return {
			program: parts[0] || fallbackProgram,
			height: parts[1] || fallbackHeight,
			width: parts[2] || fallbackWidth
		};
	}

	function ensureDetailsNode() {
		var root = outRoot();
		var doc = outDoc();
		var details = elementChildren(root, "Details")[0];
		if (!details && root && doc) {
			details = doc.createElement("Details");
			root.appendChild(details);
		}
		return details;
	}

	function checkAvailability(itemCode) {
		var exists = false;
		elementChildren(ensureDetailsNode(), "ItemDetail").forEach(function (item) {
			if (getAttr(item, "ItemCode") === trim(itemCode)) {
				exists = true;
			}
		});
		return exists;
	}

	function itemDetailFromSource(source, refNo) {
		var doc = outDoc();
		var item = doc.createElement("ItemDetail");
		setAttr(item, "ItemCode", getAttr(source, "ItemCode"));
		setAttr(item, "CLACODE", getAttr(source, "ClassCode"));
		setAttr(item, "QTY", "");
		setAttr(item, "MRSNO", "N");
		setAttr(item, "ISSNO", "N");
		setAttr(item, "ENTRYNO", getAttr(source, "EntryNo") || String(elementChildren(ensureDetailsNode(), "ItemDetail").length + 1));
		setAttr(item, "UNIT", field("hOrgID") ? field("hOrgID").value : "");
		setAttr(item, "ITEMNAME", getAttr(source, "ItemName"));
		setAttr(item, "UOM", getAttr(source, "StoresUoM"));
		setAttr(item, "ATTRIBUTELIST", getAttr(source, "AttributeList") === "0" ? "" : getAttr(source, "AttributeList"));
		setAttr(item, "RefNo", refNo || "");
		setAttr(item, "ReqQty", getAttr(source, "Qty") || "0");
		setAttr(item, "RECEIPTNUM", getAttr(source, "ReceiptNum"));
		setAttr(item, "BYPRODUCT", "P");
		return item;
	}

	function appendItemsFromRoot(root, refNo) {
		var details = ensureDetailsNode();
		elementChildren(root, "Item").forEach(function (source) {
			if (!checkAvailability(getAttr(source, "ItemCode"))) {
				details.appendChild(itemDetailFromSource(source, refNo));
			}
		});
	}

	function saveReceiptLotData(edit) {
		syncPost(edit ? "XMLSave.asp?Name=ReceiptLotDataEdit&SessionFlag=true" : "XMLSave.asp?Name=ReceiptLotData&SessionFlag=true", serializeXml(xmlIsland("OutData2")));
	}

	function table() {
		return document.getElementById("tblData");
	}

	function clearTableRows() {
		var tbl = table();
		while (tbl && tbl.rows.length > 2) {
			tbl.deleteRow(2);
		}
		rowCount = 0;
	}

	function addCell(row, className, align) {
		var cell = row.insertCell(-1);
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		return cell;
	}

	function makeInput(name, value, className, size) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.size = size || 12;
		input.className = className || "FormElem";
		input.value = value == null ? "" : String(value);
		input.style.textAlign = "right";
		return input;
	}

	function makeIconButton(name, alt, handler) {
		var link = document.createElement("a");
		var img = document.createElement("img");
		link.href = "javascript:void(0)";
		img.name = name;
		img.border = "0";
		img.src = "../../assets/images/iTMS%20Icons/Entry.gif";
		img.width = 15;
		img.height = 15;
		img.alt = alt;
		img.onclick = handler;
		link.appendChild(img);
		return link;
	}

	function attributeCaption(attributeList) {
		var caption = "";
		var parts = trim(attributeList).split(",");
		parts.forEach(function (part) {
			var valueParts;
			var attrParts;
			if (!trim(part)) {
				return;
			}
			valueParts = part.split(":");
			if (valueParts.length === 1) {
				try {
					caption = trim(syncGet("../../include/GetAttrName.asp?AttID=" + encodeURIComponent(valueParts[0])).responseText) || caption;
				} catch (ignoreAttrName) {}
				return;
			}
			attrParts = (valueParts[0] || "").split("#");
			if (trim(attrParts[1]) && trim(attrParts[1]) !== "0") {
				caption += "[" + valueParts[1] + "]";
			}
		});
		return caption;
	}

	function ensureStorage(itemNode, itemCode) {
		var xhr = syncGet("GetItemStoreInfo.asp?ItemCode=" + encodeURIComponent(itemCode));
		var storeName = "";
		var storeCount = 0;
		var root;
		var hasStorage = false;
		if (loadXmlResponse("NewData", xhr)) {
			root = xmlRoot(xmlIsland("NewData"));
			storeCount = toNumber(getAttr(root, "StoreCount"));
			elementChildren(itemNode, "STORAGE").forEach(function () {
				hasStorage = true;
			});
			elementChildren(root, "Storage").forEach(function (storage) {
				var bin = getAttr(storage, "BinNo");
				storeName = getAttr(storage, "StoreName");
				if (!hasStorage) {
					var node = outDoc().createElement("STORAGE");
					setAttr(node, "STORE", getAttr(storage, "LocNo"));
					setAttr(node, "BIN", !trim(bin) || bin === "0" ? "NULL" : bin);
					setAttr(node, "APPLICABLE", "IN");
					setAttr(node, "MONTHYEAR", "");
					setAttr(node, "QTY", "0");
					setAttr(node, "STORAGEVALUE", "0");
					setAttr(node, "SQ", "0");
					itemNode.appendChild(node);
					hasStorage = true;
				}
			});
		}
		return {
			name: storeName,
			count: storeCount
		};
	}

	function matchingItem(itemCode, classCode, entryNo) {
		var found = null;
		descendants(outRoot(), "ItemDetail").some(function (item) {
			if (getAttr(item, "ItemCode") === trim(itemCode) && getAttr(item, "CLACODE") === trim(classCode) && getAttr(item, "ENTRYNO") === trim(entryNo)) {
				found = item;
				return true;
			}
			return false;
		});
		return found;
	}

	function firstStorage(itemNode) {
		return elementChildren(itemNode, "STORAGE")[0] || null;
	}

	function displayItemName(itemNode) {
		var itemName = getAttr(itemNode, "ITEMNAME").replace(/~~/g, "\"");
		var attrs = attributeCaption(getAttr(itemNode, "ATTRIBUTELIST"));
		return attrs ? itemName + " [ " + attrs + " ]" : itemName;
	}

	window.AddItem = function () {
		var size = popupSize("1", "ItemSelectRelPartyCommon.asp", "500", "750");
		var orgId = field("hOrgID") ? field("hOrgID").value : "";
		var itemType = field("hIType") ? field("hIType").value : "";
		var url = "../../Common/" + size.program + "?orgID=" + encodeURIComponent(orgId) + "&sIType=" + encodeURIComponent(itemType) + "&Stock=N&hSelectMode=M&Flag=1&hDispButt=Y&CallFrom=PUR";
		openModal(url, xmlIsland("OutData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function () {
			appendItemsFromRoot(xmlRoot(xmlIsland("OutData")), "");
			window.DisplayTable(new Date());
			saveReceiptLotData(false);
		});
		return false;
	};

	window.CheckAvailability = checkAvailability;

	window.DisplayItem = function (name) {
		var parts = String(name || "").split("A");
		var tempValues = parts[0] + "A" + parts[2] + "A" + parts[1] + "A" + parts[3];
		openModal("itmDetailsPop.asp?sTemp=" + encodeURIComponent(tempValues), "", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No");
		return false;
	};

	window.CheckStorage = function (obj) {
		var parts = String(obj && obj.name || "").split("Z");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var orgCode = parts[4] || "";
		openModal("StorageSelectForItemPop.asp?sUnit=" + encodeURIComponent(orgCode) + "&ItemCode=" + encodeURIComponent(itemCode), xmlIsland("StoreData"), "dialogHeight:200px;dialogWidth:300px;", function () {
			var storeRoot = xmlRoot(xmlIsland("StoreData"));
			var item = matchingItem(itemCode, classCode, entryNo);
			var storeField = field("txtStore" + classCode + "A" + itemCode + "A" + entryNo);
			elementChildren(storeRoot, "STOREDET").some(function (store) {
				var unitStore = getAttr(store, "UNITSTORE").split("-");
				var node;
				if (!item) {
					return true;
				}
				node = outDoc().createElement("STORAGE");
				setAttr(node, "STORE", trim(unitStore[0]));
				setAttr(node, "BIN", !trim(unitStore[1]) || unitStore[1] === "0" ? "NULL" : trim(unitStore[1]));
				setAttr(node, "APPLICABLE", "IN");
				setAttr(node, "MONTHYEAR", "");
				setAttr(node, "QTY", "0");
				setAttr(node, "STORAGEVALUE", "0");
				setAttr(node, "SQ", "0");
				item.appendChild(node);
				if (storeField) {
					storeField.value = getAttr(store, "STORE");
				}
				return true;
			});
			saveReceiptLotData(false);
		});
		return false;
	};

	function processSelectedReferences(refType) {
		var ndRoot = xmlRoot(xmlIsland("OutData"));
		var refCodes = [];
		var refDates = [];
		var refText = [];
		elementChildren(ndRoot, "Reference").forEach(function (node) {
			refCodes.push(getAttr(node, "ReferenceNo"));
			refDates.push(getAttr(node, "ReferenceDate"));
			refText.push(getAttr(node, "ReferenceCode") + " - " + getAttr(node, "ReferenceDate"));
		});
		setText("RefNoDate", refText.join(","));
		if (field("hRefNo")) {
			field("hRefNo").value = refCodes.join(",");
		}
		if (field("hRefDate")) {
			field("hRefDate").value = refDates.join(",");
		}
		if (!refCodes.length) {
			return;
		}
		function loadReferenceItems() {
			var xhr = syncGet("InvGetItemDetForInternalReceipt.asp?RefType=" + encodeURIComponent(refType) + "&RefCodes=" + encodeURIComponent(refCodes.join(",")) + "&orgID=" + encodeURIComponent(field("hOrgID") ? field("hOrgID").value : ""));
			if (!loadXmlResponse("ItemData", xhr)) {
				alert(trim(xhr.responseText));
				return;
			}
			if (field("hItemType") && !trim(field("hItemType").value)) {
				field("hItemType").value = getAttr(xmlRoot(xmlIsland("ItemData")), "ItemType");
			}
			appendItemsFromRoot(xmlRoot(xmlIsland("ItemData")), refCodes.join(","));
		}
		if (trim(refType) === "12") {
			var ret = trim(syncGet("GetIssItemReturnable.asp?RefCodes=" + encodeURIComponent(refCodes.join(","))).responseText);
			var retParts = ret.split(":");
			if (retParts.length === 2 && retParts[0] === "Y" && retParts[1] === "D") {
				var size = popupSize("1", "ItemSelectRelPartyCommon.asp", "500", "750");
				openModal("../../Common/" + size.program + "?orgID=" + encodeURIComponent(field("hOrgID").value) + "&sIType=" + encodeURIComponent(field("hIType") ? field("hIType").value : "") + "&hSelectMode=M&Flag=1", xmlIsland("ItemData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function () {
					appendItemsFromRoot(xmlRoot(xmlIsland("ItemData")), refCodes.join(","));
				});
				return;
			}
			if (ret && retParts.length !== 2) {
				alert(ret);
			}
		}
		loadReferenceItems();
	}

	window.GetDetails = function () {
		var refType = selectedValue("selRefName");
		var orgId = field("hOrgID") ? field("hOrgID").value : "";
		if (field("selDepart") && field("selDepart").selectedIndex === 0) {
			alert("Select Received From");
			field("selDepart").focus();
			return false;
		}
		ensureDetailsNode();
		window.sUnit = orgId;
		window.sIType = field("hIType") ? field("hIType").value : "";
		if (typeof window.RefTypeSelection === "function") {
			window.RefTypeSelection(refType, orgId, "", "N", 1, "Y", 0, "PUR", function () {
				if (trim(refType) !== "N") {
					processSelectedReferences(refType);
				} else {
					appendItemsFromRoot(xmlRoot(xmlIsland("OutData")), "");
				}
				window.DisplayTable(new Date());
				saveReceiptLotData(false);
			});
		}
		return false;
	};

	window.DisplayTable = function () {
		var tbl = table();
		var mode = field("hMode") ? field("hMode").value : "N";
		clearTableRows();
		descendants(outRoot(), "ItemDetail").forEach(function (itemNode) {
			var entryNo = getAttr(itemNode, "ENTRYNO");
			var itemCode = getAttr(itemNode, "ItemCode");
			var classCode = getAttr(itemNode, "CLACODE");
			var orgId = getAttr(itemNode, "UNIT");
			var storeInfo = ensureStorage(itemNode, itemCode);
			var qty = mode === "E" ? getAttr(itemNode, "QTY") : "0";
			var itemRate = getAttr(itemNode, "ITEMRATE");
			var value = mode === "E" ? String(toNumber(qty) * (itemRate ? toNumber(itemRate) : 1)) : "0";
			var row = tbl.insertRow(-1);
			var cell;
			var link;
			var input;
			var checkbox;
			rowCount += 1;
			addCell(row, "ExcelSerial", "center").textContent = String(rowCount);
			cell = addCell(row, "ExcelDisplayCell", "left");
			link = document.createElement("a");
			link.href = "javascript:void(0)";
			link.className = "ExcelDisplayLink";
			link.name = "lnkA" + itemCode + "A" + classCode + "A" + orgId;
			link.textContent = displayItemName(itemNode);
			link.onclick = function () {
				return window.DisplayItem(this.name);
			};
			cell.appendChild(link);
			cell = addCell(row, "ExcelDisplayCell", "left");
			input = makeInput("txtStore" + classCode + "A" + itemCode + "A" + entryNo, storeInfo.name, "FormElemRead", 18);
			input.readOnly = true;
			cell.appendChild(input);
			if (storeInfo.count > 1) {
				cell.appendChild(makeIconButton("btnZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z" + orgId + "Z" + getAttr(itemNode, "ATTRIBUTELIST"), "Pick Details", function () {
					return window.CheckStorage(this);
				}));
			}
			cell = addCell(row, "ExcelInputCell");
			input = makeInput("txtQTY" + classCode + "A" + itemCode + "A" + entryNo, qty, "FormElem", 12);
			input.setAttribute("onkeypress", "DoKeyPress('',7,3)");
			cell.appendChild(input);
			cell = addCell(row, "ExcelInputCell");
			input = makeInput("txtVAL" + classCode + "A" + itemCode + "A" + entryNo, value, "FormElem", 12);
			input.setAttribute("onkeypress", "DoKeyPress('Y',10,2)");
			cell.appendChild(input);
			cell = addCell(row, "ExcelFieldCell", "center");
			input = document.createElement("input");
			input.type = "button";
			input.value = " Yes ";
			input.name = "btn:" + classCode + ":" + itemCode + ":" + entryNo;
			input.className = "AddButtonX";
			input.onclick = function () {
				return window.GetLot(this);
			};
			cell.appendChild(input);
			cell = addCell(row, "ExcelFieldCell", "center");
			checkbox = document.createElement("input");
			checkbox.type = "checkbox";
			checkbox.value = " W ";
			checkbox.name = "chkProduct" + classCode + "A" + itemCode + "A" + entryNo;
			checkbox.className = "FormElem";
			checkbox.checked = getAttr(itemNode, "BYPRODUCT") === "W";
			cell.appendChild(checkbox);
		});
		if (field("hCtr")) {
			field("hCtr").value = tbl ? String(tbl.rows.length + 1) : "0";
		}
	};

	window.DoChanges = function (obj) {
		if (!obj || obj.value !== "PRD") {
			return false;
		}
		openModal("../../Common/WorkCenterPopup.asp", "", "dialogHeight:150px;dialogWidth:300px;", function (workCenter) {
			var nodes = elementChildren(workCenter);
			if (nodes.length) {
				if (field("hWCCode")) {
					field("hWCCode").value = getAttr(nodes[0], "Code");
				}
				setText("spanWCName", getAttr(nodes[0], "Name"));
			} else if (!confirm("Do you want to continue without selecting the workcenter?")) {
				window.DoChanges(obj);
			}
		});
		return false;
	};

	window.DoDisable = function () {
		return false;
	};

	window.checkNumbers = checkNumbers;
	window.ClearTable = clearTableRows;

	window.clearXML = function () {
		var root = xmlRoot(xmlIsland("OutData"));
		elementChildren(root, "STORAGE").forEach(function (node) {
			root.removeChild(node);
		});
	};

	window.GetLot = function (obj) {
		var parts = String(obj && obj.name || "").split(":");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var qty = field("txtQTY" + classCode + "A" + itemCode + "A" + entryNo);
		var value = field("txtVAL" + classCode + "A" + itemCode + "A" + entryNo);
		var storeName = field("txtStore" + classCode + "A" + itemCode + "A" + entryNo);
		var item = matchingItem(itemCode, classCode, entryNo);
		var storage = firstStorage(item);
		var receiptType = getAttr(item, "RECEIPTNUM");
		var receivedOn = getDatePickerValue("ctlRcvdOn");
		var loc;
		var bin;
		var tempValues;
		if (!qty || trim(qty.value) === "") {
			alert("Enter Quantity");
			if (qty) {
				qty.select();
			}
			return false;
		}
		if (!checkNumbers(qty.value)) {
			alert("Enter Numerals Only");
			qty.select();
			return false;
		}
		if (!value || trim(value.value) === "") {
			alert("Enter Value");
			if (value) {
				value.select();
			}
			return false;
		}
		if (!checkNumbers(value.value)) {
			alert("Enter Numerals Only");
			value.select();
			return false;
		}
		if (!storage) {
			alert("Select Storage Location and continue");
			return false;
		}
		loc = getAttr(storage, "STORE");
		bin = getAttr(storage, "BIN");
		if (!trim(loc)) {
			alert("Select Storage Location and continue");
			return false;
		}
		if (trim(receiptType) === "N") {
			return false;
		}
		tempValues = [
			receiptType,
			itemCode,
			classCode,
			field("hOrgID") ? field("hOrgID").value : "",
			loc,
			bin,
			qty.value,
			(storeName ? storeName.value : "").replace(/&/g, "and"),
			field("hStoresUom") ? field("hStoresUom").value : "",
			value.value,
			receivedOn,
			field("hIType") ? field("hIType").value : "",
			getAttr(item, "ATTRIBUTELIST")
		].join("``");
		openModal("../../Common/PackingLotSerialDetails.asp?sTemp=" + encodeURIComponent(tempValues) + "&CallFrom=IRCPT&RefNo=" + encodeURIComponent(field("hRcptNo") ? field("hRcptNo").value : ""), xmlIsland("OutData2"), "dialogHeight:580px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function () {
			var updated = firstStorage(item);
			if (updated) {
				setAttr(updated, "MONTHYEAR", receivedOn);
				if (qty) {
					qty.value = getAttr(updated, "QTY");
				}
				if (value) {
					value.value = getAttr(updated, "STORAGEVALUE");
				}
			}
		});
		return false;
	};

	function updateXmlFromRows() {
		descendants(outRoot(), "ItemDetail").forEach(function (item) {
			var itemCode = getAttr(item, "ItemCode");
			var classCode = getAttr(item, "CLACODE");
			var entryNo = getAttr(item, "ENTRYNO");
			var qty = field("txtQTY" + classCode + "A" + itemCode + "A" + entryNo);
			var value = field("txtVAL" + classCode + "A" + itemCode + "A" + entryNo);
			var product = field("chkProduct" + classCode + "A" + itemCode + "A" + entryNo);
			var totalQty = 0;
			elementChildren(item, "STORAGE").forEach(function (storage) {
				setAttr(storage, "MONTHYEAR", getDatePickerValue("ctlRcvdOn"));
				setAttr(storage, "QTY", trim(qty && qty.value));
				setAttr(storage, "STORAGEVALUE", trim(value && value.value));
				setAttr(storage, "SQ", "0");
				totalQty += toNumber(qty && qty.value);
			});
			setAttr(item, "QTY", totalQty);
			setAttr(item, "BYPRODUCT", product && product.checked ? "W" : "P");
		});
	}

	function validateRows() {
		var controls = form() && form().elements || [];
		for (var i = 0; i < controls.length; i += 1) {
			if (controls[i].type === "text" && String(controls[i].name || "").indexOf("txtQTY") === 0) {
				if (!trim(controls[i].value) || trim(controls[i].value) === "0") {
					alert("Enter Quantity");
					controls[i].select();
					return false;
				}
				if (!checkNumbers(controls[i].value)) {
					alert("Enter Numerals Only");
					controls[i].select();
					return false;
				}
				if (!controls[i + 1] || !trim(controls[i + 1].value) || trim(controls[i + 1].value) === "0") {
					alert("Enter Value");
					if (controls[i + 1]) {
						controls[i + 1].select();
					}
					return false;
				}
				if (!checkNumbers(controls[i + 1].value)) {
					alert("Enter Numerals Only");
					controls[i + 1].select();
					return false;
				}
			}
		}
		return true;
	}

	function validateStorageAndLots() {
		var ok = true;
		descendants(outRoot(), "ItemDetail").some(function (item) {
			if (!elementChildren(item, "STORAGE").length) {
				alert("Select Storage Location for " + getAttr(item, "ITEMNAME"));
				ok = false;
				return true;
			}
			elementChildren(item, "STORAGE").some(function (storage) {
				if (trim(getAttr(item, "RECEIPTNUM")) !== "N" && !elementChildren(storage).length) {
					alert("Enter Lot / Serial Details for " + getAttr(item, "ITEMNAME"));
					ok = false;
					return true;
				}
				return false;
			});
			return !ok;
		});
		return ok;
	}

	window.CheckSubmit = function (finFrom, finTo) {
		var root = outRoot();
		var refName = selectedValue("selRefName");
		var refType = refName === "12" ? "I" : refName === "17" ? "P" : "N";
		var mode;
		if (rowCount > 0 && !validateRows()) {
			return false;
		}
		updateXmlFromRows();
		if (!validateStorageAndLots()) {
			return false;
		}
		setAttr(root, "DEPT", selectedValue("selDepart"));
		setAttr(root, "PACKNUM", field("hDept") && field("hDept").value === "PRD" && field("hType") && field("hType").value === "T" ? "N" : "");
		setAttr(root, "SRCREFTYPE", refType);
		setAttr(root, "SRCREFNO", field("hRefNo") ? field("hRefNo").value : "");
		setAttr(root, "STYPE", field("hWCCode") ? field("hWCCode").value : "");
		setAttr(root, "APPREFTYPE", refName);
		setAttr(root, "APPREFNO", field("hRefNo") ? field("hRefNo").value : "");
		setAttr(root, "APPREFDATE", field("hRefDate") ? field("hRefDate").value : "");
		setAttr(root, "RCVDON", getDatePickerValue("ctlRcvdOn"));
		setAttr(root, "AUTOACCOUNT", field("hAutoAccount") ? field("hAutoAccount").value : "");
		mode = field("hMode") ? field("hMode").value : "N";
		syncPost(mode === "E" ? "../Master/XMLSave.asp?SessionFlag=true&Value=ReceiptLotDataEdit&Folder=Transaction" : "../Master/XMLSave.asp?SessionFlag=true&Value=ReceiptLotData&Folder=Transaction", serializeXml(xmlIsland("OutData2")));
		form().action = mode === "E" ? "InternalReceiptUpdate.asp?CurrDate=" + encodeURIComponent(field("hCurrDate").value) + "&RcptNo=" + encodeURIComponent(field("hRcptNo").value) : "receiptNewInsert.asp?CurrDate=" + encodeURIComponent(field("hCurrDate").value);
		form().submit();
		return false;
	};

	window.SelectReference = function (refType) {
		var orgId = field("hOrgID") ? field("hOrgID").value : "";
		if (refType !== "I") {
			return false;
		}
		openModal("IntRcptRefNoSel.asp?RefType=" + encodeURIComponent(refType) + "&OrgID=" + encodeURIComponent(orgId), "", "dialogHeight=200px;dialogWidth:500px;", function (rootRef) {
			var nodes = elementChildren(rootRef);
			if (nodes.length && field("txtSource")) {
				field("txtSource").value = getAttr(nodes[0], "value");
				setText("spanSource", getAttr(nodes[0], "name"));
			}
		});
		return false;
	};

	window.EditInit = function (receiptNo) {
		var mode = field("hMode") ? field("hMode").value : "N";
		var xhr;
		var root;
		var refXml;
		if (trim(mode) !== "E") {
			return false;
		}
		xhr = syncGet("PopIntRcptData.asp?RcptNo=" + encodeURIComponent(receiptNo));
		if (loadXmlResponse("OutData2", xhr)) {
			root = outRoot();
			setDatePickerValue("ctlRcvdOn", getAttr(root, "RCVDON"));
			selectByValue("selDepart", getAttr(root, "DEPT"));
			selectByValue("selRefName", getAttr(root, "APPREFTYPE"));
			xhr = syncGet("../../Common/GetInfoForRefType.asp?orgID=" + encodeURIComponent(getAttr(root, "ORGCODE")) + "&RefType=" + encodeURIComponent(getAttr(root, "APPREFTYPE")) + "&RefNo=" + encodeURIComponent(getAttr(root, "APPREFNO")));
			if (loadXmlResponse("RefXML", xhr)) {
				refXml = xmlRoot(xmlIsland("RefXML"));
				elementChildren(refXml, "Ref").some(function (ref) {
					setText("RefNoDate", getAttr(ref, "Code") + " - " + getAttr(ref, "Date"));
					return true;
				});
			}
			window.DisplayTable(new Date());
			saveReceiptLotData(true);
		}
		return false;
	};

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseDate(value) {
		var text = trim(value);
		var match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		var year;
		var date;
		if (!text) {
			return null;
		}
		if (match) {
			year = Number(match[3]);
			if (year < 100) {
				year += 2000;
			}
			date = new Date(year, Number(match[2]) - 1, Number(match[1]));
			return isNaN(date.getTime()) ? null : date;
		}
		date = new Date(text);
		return isNaN(date.getTime()) ? null : date;
	}

	function todayText() {
		var date = new Date();
		return pad2(date.getDate()) + "/" + pad2(date.getMonth() + 1) + "/" + date.getFullYear();
	}

	function getDatePickerValue(name) {
		var control = field(name);
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		return control.value || "";
	}

	function setDatePickerValue(name, value) {
		var control = field(name);
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

	window.setdate = function () {
		var fromDate = field("hMinDate") ? field("hMinDate").value : "";
		var toDate = field("hMaxDate") ? field("hMaxDate").value : "";
		var max = parseDate(toDate) && parseDate(toDate) < new Date() ? toDate : todayText();
		var control = field("ctlRcvdOn");
		if (control) {
			if (typeof control.SetMinDate === "function") {
				control.SetMinDate(fromDate);
			} else if (typeof control.setMinDate === "function") {
				control.setMinDate(fromDate);
			}
			if (typeof control.SetMaxDate === "function") {
				control.SetMaxDate(max);
			} else if (typeof control.setMaxDate === "function") {
				control.setMaxDate(max);
			}
			if (parseDate(toDate) && parseDate(toDate) < new Date()) {
				setDatePickerValue("ctlRcvdOn", toDate);
			}
		}
		return false;
	};
}(window, document));
