(function (window, document) {
	"use strict";

	var rows = [];

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
		var value = frm && frm.elements ? frm.elements[name] : null;
		if (value && value.length && !value.tagName) {
			return value[0] || null;
		}
		return value || null;
	}

	function byId(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function today() {
		var value = field("hCreatedOn") && field("hCreatedOn").value;
		if (value) {
			return value;
		}
		var date = new Date();
		var day = date.getDate() < 10 ? "0" + date.getDate() : String(date.getDate());
		var month = date.getMonth() + 1 < 10 ? "0" + (date.getMonth() + 1) : String(date.getMonth() + 1);
		return day + "/" + month + "/" + date.getFullYear();
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name);
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
		if (value && typeof value.xml === "string") {
			return value.xml;
		}
		if (doc && typeof doc.xml === "string") {
			return doc.xml;
		}
		return new XMLSerializer().serializeToString(doc || root);
	}

	function loadXmlIsland(name, text) {
		var island = xmlIsland(name);
		if (island && typeof island.loadXML === "function") {
			island.loadXML(text || "<Root/>");
			return xmlRoot(island);
		}
		return new DOMParser().parseFromString(text || "<Root/>", "text/xml").documentElement;
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

	function removeChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function createElement(doc, name) {
		return (doc || document).createElement(name);
	}

	function importNode(doc, node) {
		return doc && doc.importNode ? doc.importNode(node, true) : node.cloneNode(true);
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

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function buildQuery(params) {
		var parts = [];
		Object.keys(params).forEach(function (key) {
			parts.push(encodeURIComponent(key) + "=" + encodeURIComponent(params[key] == null ? "" : params[key]));
		});
		return parts.join("&");
	}

	function clearTable(tableId, keepRows) {
		var table = byId(tableId);
		if (!table) {
			return;
		}
		while (table.rows.length > keepRows) {
			table.deleteRow(keepRows);
		}
	}

	function clearXml(name) {
		removeChildren(xmlRoot(xmlIsland(name)));
	}

	function appendCell(row, className, align, content) {
		var cell = row.insertCell(row.cells.length);
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		if (content == null) {
			return cell;
		}
		if (typeof content === "string" || typeof content === "number") {
			cell.textContent = String(content);
		} else {
			cell.appendChild(content);
		}
		return cell;
	}

	function makeHidden(name, value) {
		var input = document.createElement("input");
		input.type = "hidden";
		input.name = name;
		input.value = value == null ? "" : String(value);
		return input;
	}

	function makeInput(name, value, readOnly) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.size = 8;
		input.value = value == null ? "" : String(value);
		input.className = readOnly ? "FormElemRead" : "FormElem";
		input.readOnly = !!readOnly;
		input.style.textAlign = "right";
		return input;
	}

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.textContent = value == null ? "" : String(value);
			element.innerText = element.textContent;
		}
	}

	function textOf(id) {
		var element = byId(id);
		return trim(element ? element.textContent || element.innerText : "");
	}

	function getItemType(itemCode) {
		var xhr = syncGet("../../include/GetItemTypeForItem.asp?ItemCode=" + encodeURIComponent(itemCode));
		var root;
		if (xhr.status >= 400 || !trim(xhr.responseText)) {
			return "";
		}
		root = loadXmlIsland("ItemTypeData", xhr.responseText);
		return getAttr(root, "ItemType");
	}

	function allItems(root) {
		return elementChildren(root, "Item");
	}

	function stockForLocation(loc) {
		var total = 0;
		elementChildren(loc, "PICK").forEach(function (pick) {
			total += toNumber(getAttr(pick, "QTYSTK"));
		});
		return total;
	}

	function tempRoot() {
		return xmlRoot(xmlIsland("TempItemData"));
	}

	function addTempItem(row, item, loc) {
		var doc = xmlDocument(xmlIsland("TempItemData")) || tempRoot().ownerDocument;
		var node = createElement(doc, "Item");
		setAttr(node, "ENTRYNO", row.index);
		setAttr(node, "ITEMCODE", getAttr(item, "ICode"));
		setAttr(node, "CLASSCODE", getAttr(item, "CCode"));
		setAttr(node, "LOCNO", getAttr(loc, "LOC"));
		setAttr(node, "BINNO", getAttr(loc, "BIN"));
		setAttr(node, "ITEMNAME", getAttr(item, "IName"));
		tempRoot().appendChild(node);
	}

	function findTempItem(row) {
		var items = elementChildren(tempRoot(), "Item");
		for (var i = 0; i < items.length; i += 1) {
			if (getAttr(items[i], "ENTRYNO") === String(row.index) &&
					getAttr(items[i], "ITEMCODE") === row.itemCode &&
					getAttr(items[i], "CLASSCODE") === row.classCode &&
					getAttr(items[i], "LOCNO") === row.locNo &&
					getAttr(items[i], "BINNO") === row.binNo) {
				return items[i];
			}
		}
		return null;
	}

	function selectedReceiptNumbering() {
		return field("hReceiptNum") ? field("hReceiptNum").value : "";
	}

	function renderSourceRow(item, loc, itemEntryNo) {
		var table = byId("tblData");
		var rowInfo;
		var row;
		var storeCell;
		var stockCell;
		var mergeCell;
		var stock;
		var input;
		var recNum = getAttr(loc, "RECNUM");
		if (!table) {
			return;
		}
		stock = stockForLocation(loc);
		rowInfo = {
			index: rows.length + 1,
			itemEntryNo: itemEntryNo,
			itemCode: getAttr(item, "ICode"),
			classCode: getAttr(item, "CCode"),
			itemName: getAttr(item, "IName"),
			locNo: getAttr(loc, "LOC"),
			binNo: getAttr(loc, "BIN"),
			locName: getAttr(loc, "LOCNAME"),
			recNum: recNum,
			value: getAttr(loc, "VALUE"),
			stock: stock,
			input: null
		};
		rows.push(rowInfo);
		addTempItem(rowInfo, item, loc);
		row = table.insertRow(table.rows.length);
		appendCell(row, "ExcelSerial", "center", rowInfo.index);
		appendCell(row, "ExcelDisplayCell", null, rowInfo.itemName);
		storeCell = appendCell(row, "ExcelDisplayCell", null, rowInfo.locName);
		storeCell.appendChild(makeHidden("hItemEntryNo" + rowInfo.index, itemEntryNo));
		storeCell.appendChild(makeHidden("hFromItemCode" + rowInfo.index, rowInfo.itemCode));
		storeCell.appendChild(makeHidden("hFromClassCode" + rowInfo.index, rowInfo.classCode));
		storeCell.appendChild(makeHidden("hFromItemName" + rowInfo.index, rowInfo.itemName));
		storeCell.appendChild(makeHidden("hFromLocNo" + rowInfo.index, rowInfo.locNo));
		storeCell.appendChild(makeHidden("hFromBinNo" + rowInfo.index, rowInfo.binNo));
		storeCell.appendChild(makeHidden("hFromItemValue" + rowInfo.index, rowInfo.value));
		stockCell = appendCell(row, "ExcelDisplayCell", "right");
		stockCell.innerHTML = "";
		var link = document.createElement("a");
		link.href = "#";
		link.className = "ExcelDisplayLink";
		link.textContent = String(stock);
		link.onclick = function () {
			return window.ViewLotDetails(rowInfo.index);
		};
		stockCell.appendChild(link);
		mergeCell = appendCell(row, "ExcelDisplayCell", "right");
		if (recNum === "LS" || recNum === "S") {
			var icon = document.createElement("img");
			icon.src = "../../assets/images/iTMS%20Icons/Entryicon.gif";
			icon.border = 0;
			icon.style.cursor = "pointer";
			icon.alt = "Pick";
			icon.onclick = function () {
				return window.GetLotDetails(rowInfo.index, rowInfo.itemCode, rowInfo.classCode, rowInfo.locNo, rowInfo.binNo);
			};
			mergeCell.appendChild(icon);
			input = makeInput("hItemStock" + rowInfo.index, "0", true);
		} else {
			input = makeInput("hItemStock" + rowInfo.index, stock, false);
			input.onchange = function () {
				window.CalculateVal(rowInfo.index);
			};
		}
		rowInfo.input = input;
		mergeCell.appendChild(input);
		if (field("hFromRcptNum")) {
			field("hFromRcptNum").value = recNum;
		}
	}

	function renderTargetRow() {
		var table = byId("tblToItemData");
		var row;
		var itemCell;
		var icon;
		if (!table) {
			return;
		}
		clearTable("tblToItemData", 1);
		row = table.insertRow(table.rows.length);
		appendCell(row, "ExcelSerial", "center", "1");
		itemCell = appendCell(row, "ExcelDisplayCell");
		var span = document.createElement("span");
		span.id = "spaMergeItem";
		itemCell.appendChild(span);
		icon = document.createElement("img");
		icon.border = 0;
		icon.src = "../../assets/images/iTMS%20icons/Entryicon.gif";
		icon.width = 12;
		icon.height = 12;
		icon.alt = "Select";
		icon.style.cursor = "pointer";
		icon.onclick = function () {
			return window.SelectItem();
		};
		itemCell.appendChild(icon);
		itemCell.appendChild(makeHidden("hMergeItemCode", ""));
		itemCell.appendChild(makeHidden("hMergeClassCode", ""));
		itemCell.appendChild(makeHidden("hLocCode", ""));
		itemCell.appendChild(makeHidden("hBinNo", ""));
		itemCell.appendChild(makeHidden("hAttID", ""));
		appendCell(row, "ExcelDisplayCell").innerHTML = "<span id=\"spaMergeStore\"></span>";
		appendCell(row, "ExcelDisplayCell", "right").innerHTML = "<span id=\"spaMergeStock\" style=\"text-align:right;\"></span>";
		appendCell(row, "ExcelDisplayCell", "right").innerHTML = "<span id=\"spaMergedQty\" style=\"text-align:right;\"></span>";
	}

	function previousMergeQuantity() {
		var total = 0;
		rows.forEach(function (row) {
			total += toNumber(row.input && row.input.value);
		});
		return total;
	}

	function firstChildElement(root) {
		var children = elementChildren(root);
		return children[0] || null;
	}

	function processSelectedMergeItem() {
		var root = xmlRoot(xmlIsland("ItemSelectData"));
		var selected = firstChildElement(root);
		var attrList;
		var attrParts;
		var optParts;
		var itemName;
		var itemCode;
		var classCode;
		var receiptNum;
		var attrId = "";
		var xhr;
		var stockRoot;
		var stock = 0;
		var locName = "";
		var locCode = "";
		var binNo = "";
		if (!selected) {
			return;
		}
		itemName = getAttr(selected, "ItemName");
		itemCode = getAttr(selected, "ItemCode");
		classCode = getAttr(selected, "ClassCode");
		receiptNum = getAttr(selected, "ReceiptNum");
		attrList = getAttr(selected, "AttributeList");
		if (attrList) {
			attrParts = attrList.split("#");
			if (attrParts.length === 2) {
				optParts = attrParts[1].split(":");
				attrId = optParts[0] || "";
			}
		}
		if (field("hReceiptNum")) {
			field("hReceiptNum").value = receiptNum;
		}
		xhr = syncGet("ItmMergeStockDet.asp?ItemCode=" + encodeURIComponent(itemCode) + "&ClassCode=" + encodeURIComponent(classCode));
		if (xhr.status < 400 && trim(xhr.responseText)) {
			stockRoot = loadXmlIsland("TempData", xhr.responseText);
			elementChildren(stockRoot, "LOCDET").some(function (loc) {
				locCode = getAttr(loc, "LOC");
				binNo = getAttr(loc, "BIN");
				locName = getAttr(loc, "LOCNAME");
				elementChildren(loc, "PICK").some(function (pick) {
					stock = toNumber(getAttr(pick, "QTYSTK"));
					return true;
				});
				return true;
			});
		} else if (trim(xhr.responseText)) {
			alert(xhr.responseText);
		}
		setText("spaMergeItem", itemName);
		field("hMergeItemCode").value = itemCode;
		field("hMergeClassCode").value = classCode;
		field("hLocCode").value = locCode;
		field("hBinNo").value = binNo;
		field("hAttID").value = attrId;
		setText("spaMergeStore", locName);
		setText("spaMergeStock", stock);
		setText("spaMergedQty", previousMergeQuantity() + stock);
	}

	function resetIssueRoot(root) {
		removeChildren(root);
		setAttr(root, "ISSTYPE", "F");
		setAttr(root, "ISSTOCODE", "INV");
		setAttr(root, "ISSTOTYPE", "DEPT");
		setAttr(root, "ISSTOSUBCODE", "");
		setAttr(root, "POConfirm", "N");
		setAttr(root, "SInvConfirm", "N");
		setAttr(root, "Invoice", "A");
		setAttr(root, "GPConfirm", "N");
		setAttr(root, "ProConfirm", "N");
		setAttr(root, "MCallFrom", "N");
		setAttr(root, "RedirectTo", "");
		setAttr(root, "AppRefType", "");
		setAttr(root, "AppRefNo", "");
		setAttr(root, "AppRefDate", "");
		setAttr(root, "ConsumptionAccHead", "");
		setAttr(root, "IssueToCode", "");
		setAttr(root, "PickPackFlag", "");
		setAttr(root, "IssFrom", "IN");
		setAttr(root, "Returnable", "N");
		setAttr(root, "ReturnItem", "S");
		setAttr(root, "TYPE", "GEN");
	}

	function resetReceiptRoot(root) {
		removeChildren(root);
		setAttr(root, "DEPT", "OTH");
		setAttr(root, "SOURCE", "N");
		setAttr(root, "ORGCODE", field("hOrgID") ? field("hOrgID").value : "");
		setAttr(root, "STYPE", "N");
		setAttr(root, "ITEMTYPE", "");
		setAttr(root, "PACKNUM", "");
		setAttr(root, "SRCREFTYPE", "N");
		setAttr(root, "SRCREFNO", "");
		setAttr(root, "RCPTNUMBERINV", selectedReceiptNumbering());
		setAttr(root, "sTypeRcpt", "");
		setAttr(root, "APPREFTYPE", "");
		setAttr(root, "APPREFNO", "");
		setAttr(root, "APPREFDATE", "");
		setAttr(root, "RCVDON", today());
		setAttr(root, "AUTOACCOUNT", "Y");
	}

	function selectedRows() {
		return rows.filter(function (row) {
			return toNumber(row.input && row.input.value) > 0;
		});
	}

	function buildReceiptData() {
		var receiptIsland = xmlIsland("IntReceipt");
		var receiptDoc = xmlDocument(receiptIsland);
		var receiptRoot = xmlRoot(receiptIsland);
		var details;
		var itemDetail;
		var storage;
		var receiptNumbering = "";
		var mergeQty = previousMergeQuantity();
		var mergeStock = toNumber(textOf("spaMergeStock"));
		var totalStoreValue = 0;
		rows.forEach(function (row) {
			totalStoreValue += toNumber(row.value);
		});
		resetReceiptRoot(receiptRoot);
		details = createElement(receiptDoc, "Details");
		receiptRoot.appendChild(details);
		if (field("hMergeItemCode") && field("hMergeItemCode").value) {
			var xhr = syncGet("../../Common/GetItemRcptNumbering.asp?ItemCode=" + encodeURIComponent(field("hMergeItemCode").value));
			if (xhr.status < 400) {
				receiptNumbering = trim(xhr.responseText);
			}
		}
		itemDetail = createElement(receiptDoc, "ItemDetail");
		setAttr(itemDetail, "ItemCode", field("hMergeItemCode").value);
		setAttr(itemDetail, "CLACODE", field("hMergeClassCode").value);
		setAttr(itemDetail, "QTY", mergeQty);
		setAttr(itemDetail, "MRSNO", "N");
		setAttr(itemDetail, "ISSNO", "N");
		setAttr(itemDetail, "ENTRYNO", "1");
		setAttr(itemDetail, "UNIT", field("hOrgID") ? field("hOrgID").value : "");
		setAttr(itemDetail, "ITEMNAME", textOf("spaMergeItem"));
		setAttr(itemDetail, "UOM", "");
		setAttr(itemDetail, "ATTRIBUTELIST", field("hAttID") ? field("hAttID").value : "");
		setAttr(itemDetail, "RefNo", "");
		setAttr(itemDetail, "RefQty", "");
		setAttr(itemDetail, "RECEIPTNUM", receiptNumbering);
		setAttr(itemDetail, "BYPRODUCT", "P");
		details.appendChild(itemDetail);
		storage = createElement(receiptDoc, "STORAGE");
		setAttr(storage, "STORE", field("hLocCode") ? field("hLocCode").value : "");
		setAttr(storage, "BIN", field("hBinNo") && field("hBinNo").value !== "0" ? field("hBinNo").value : "NULL");
		setAttr(storage, "APPLICABLE", "IN");
		setAttr(storage, "MONTHYEAR", today());
		setAttr(storage, "QTY", mergeQty - mergeStock);
		setAttr(storage, "STORAGEVALUE", totalStoreValue);
		itemDetail.appendChild(storage);
		appendReceiptLotSerials(receiptDoc, storage);
	}

	function appendReceiptLotSerials(doc, storage) {
		var receiptNum = selectedReceiptNumbering();
		selectedRows().forEach(function (row) {
			var item = findTempItem(row);
			var pickRoot = item && elementChildren(item, "Pick")[0];
			if (!pickRoot) {
				return;
			}
			elementChildren(pickRoot, "PICK").forEach(function (pick) {
				var selections = elementChildren(pick, "Selection").filter(function (selection) {
					return getAttr(selection, "YesNo") === "Y";
				});
				var qty = 0;
				var lotSerial;
				if ((receiptNum === "LS" || receiptNum === "L") && !selections.length) {
					return;
				}
				if (receiptNum === "S" && !selections.length) {
					return;
				}
				selections.forEach(function (selection) {
					qty += toNumber(getAttr(selection, "Qty"));
				});
				if (!qty) {
					return;
				}
				lotSerial = createElement(doc, "LotSerial");
				setAttr(lotSerial, "QTYIN", "N");
				setAttr(lotSerial, "TARE", "0");
				setAttr(lotSerial, "LOT", getAttr(pick, "LOTNO"));
				setAttr(lotSerial, "SERIALFROM", "");
				setAttr(lotSerial, "SERIALTO", "");
				setAttr(lotSerial, "TAREWEIGHT", "U");
				setAttr(lotSerial, "IVALUE", row.value);
				setAttr(lotSerial, "QTY", qty);
				setAttr(lotSerial, "COUNTER", storage.childNodes.length + 1);
				setAttr(lotSerial, "STAGE", "select");
				setAttr(lotSerial, "ALTGROSS", "0");
				setAttr(lotSerial, "ALTNETT", "0");
				setAttr(lotSerial, "ALTUOM", "select");
				setAttr(lotSerial, "AUTOGEN", "");
				storage.appendChild(lotSerial);
				selections.forEach(function (selection) {
					appendReceiptLotSerialDetail(doc, lotSerial, selection, getAttr(pick, "LOTNO"));
				});
			});
		});
	}

	function appendReceiptLotSerialDetail(doc, lotSerial, selection, lotNo) {
		var xhr = syncGet("GetLotDetails.asp?SerialNo=" + encodeURIComponent(getAttr(selection, "SerialNo")) + "&InvNo=" + encodeURIComponent(getAttr(selection, "InvRecNo")));
		var root;
		if (xhr.status >= 400 || !trim(xhr.responseText)) {
			return;
		}
		root = loadXmlIsland("LotData", xhr.responseText);
		elementChildren(root).forEach(function (lot) {
			var detail = createElement(doc, "LotSerialDetails");
			setAttr(detail, "LOTSERIAL", getAttr(selection, "SerialNo"));
			setAttr(detail, "QTYREC", getAttr(selection, "Qty"));
			setAttr(detail, "TAREREC", "0");
			setAttr(detail, "SELLINGTYPE", getAttr(lot, "SellNo"));
			setAttr(detail, "WEIGHTSTYPE", getAttr(lot, "WeightPerSellForm"));
			setAttr(detail, "PACKINGTYPE", getAttr(lot, "PackCode"));
			setAttr(detail, "LOT", lotNo);
			setAttr(detail, "SELLINGFORM", getAttr(lot, "SellForm"));
			setAttr(detail, "PACKNUMBER", getAttr(lot, "PackNo"));
			setAttr(detail, "IVALUE", "0");
			setAttr(detail, "ATTRIBUTELIST", field("hAttID") ? field("hAttID").value : "");
			setAttr(detail, "NOOFCONE", "0");
			setAttr(detail, "SUBLEVELID", "");
			setAttr(detail, "SQ", "");
			lotSerial.appendChild(detail);
		});
	}

	function buildIssueData() {
		var issueIsland = xmlIsland("IssueData");
		var issueDoc = xmlDocument(issueIsland);
		var issueRoot = xmlRoot(issueIsland);
		var groups = {};
		var entryNo = 0;
		resetIssueRoot(issueRoot);
		selectedRows().forEach(function (row) {
			var key = row.itemCode + ":" + row.classCode;
			var group = groups[key];
			var qty = toNumber(row.input.value);
			if (!group) {
				entryNo += 1;
				group = groups[key] = {
					total: 0,
					pick: null,
					item: createElement(issueDoc, "ITEM")
				};
				setAttr(group.item, "ENTRYNO", entryNo);
				setAttr(group.item, "ITMCODE", row.itemCode);
				setAttr(group.item, "CLACODE", row.classCode);
				setAttr(group.item, "ITMNAME", row.itemName);
				setAttr(group.item, "SSTORE", "");
				setAttr(group.item, "REQQTY", "0");
				setAttr(group.item, "REQBY", "");
				setAttr(group.item, "REMARKS", "");
				setAttr(group.item, "ITEMTYPE", field("hFromItemType") ? field("hFromItemType").value : "");
				setAttr(group.item, "ISSUEDATE", today());
				setAttr(group.item, "TRAQTY", "0");
				setAttr(group.item, "PRQTY", "0");
				setAttr(group.item, "IVALUE", "0");
				setAttr(group.item, "ORGCODE", field("hOrgID") ? field("hOrgID").value : "");
				setAttr(group.item, "MRSNO", "");
				setAttr(group.item, "MRSDATE", "");
				setAttr(group.item, "ATTRIBUTELIST", "");
				setAttr(group.item, "CREATEDBY", field("hUserID") ? field("hUserID").value : "");
				setAttr(group.item, "CREATEDON", today());
				setAttr(group.item, "RETURNABLE", "N");
				setAttr(group.item, "RefNo", "");
				setAttr(group.item, "ONLYLOT", "");
				setAttr(group.item, "RETURNITEM", "S");
				setAttr(group.item, "MatType", "");
				issueRoot.appendChild(group.item);
				if (row.recNum === "N") {
					group.pick = createElement(issueDoc, "Pick");
					setAttr(group.pick, "TOT", "0");
					setAttr(group.pick, "NoofPack", "");
					group.item.appendChild(group.pick);
				}
			}
			group.total += qty;
			if (row.recNum === "N") {
				appendIssueStore(issueDoc, group.pick, row, qty);
			} else {
				appendIssuePickFromTemp(issueDoc, group.item, row);
			}
		});
		Object.keys(groups).forEach(function (key) {
			setAttr(groups[key].item, "ISSQTY", groups[key].total);
			if (groups[key].pick) {
				setAttr(groups[key].pick, "TOT", groups[key].total);
			}
		});
	}

	function appendIssueStore(doc, pick, row, qty) {
		var store = createElement(doc, "STORE");
		setAttr(store, "LOC", row.locNo);
		setAttr(store, "BIN", row.binNo);
		setAttr(store, "LOTNO", "N/A");
		setAttr(store, "INVRECNO", "");
		setAttr(store, "QTYISS", qty);
		setAttr(store, "NoofPack", "");
		pick.appendChild(store);
	}

	function appendIssuePickFromTemp(doc, issueItem, row) {
		var item = findTempItem(row);
		elementChildren(item, "Pick").forEach(function (pickRoot) {
			issueItem.appendChild(importNode(doc, pickRoot));
		});
	}

	window.FnInit = function () {
		window.GetXML();
		window.DisplayDetails();
		return false;
	};

	window.ViewLotDetails = function (index) {
		var code = field("hFromItemCode" + index) ? field("hFromItemCode" + index).value : "";
		openDialog("ViewLotDetailsPop.asp?ItemCode=" + encodeURIComponent(code), "", "dialogWidth:450px;dialogHeight:450px;Status:No;", function () {});
		return false;
	};

	window.GetLotDetails = function (index, itemCode, classCode, locNo, binNo) {
		var row = rows[toNumber(index) - 1];
		var param;
		var item;
		var pickRoot;
		if (!field("hMergeItemCode") || trim(field("hMergeItemCode").value) === "") {
			alert("Select Item Merged With");
			return false;
		}
		param = [itemCode, classCode, row ? row.itemName : "", locNo, binNo, "", "", "", index].join(":");
		openDialog("ItemMergePickPop.asp?sTemp=" + encodeURIComponent(param), xmlIsland("TempItemData"), "dialogHeight:390px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", function () {
			if (getAttr(tempRoot(), "DONE") === "YES" && row) {
				item = findTempItem(row);
				pickRoot = item && elementChildren(item, "Pick")[0];
				if (pickRoot && row.input) {
					row.input.value = getAttr(pickRoot, "TOT") || "0";
				}
				window.CalculateVal(index);
			}
		});
		return false;
	};

	window.SelectItem = function () {
		var sizeText = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup("1") : "ItemSelectRelPartyCommon.asp:500:850";
		var size = String(sizeText || "ItemSelectRelPartyCommon.asp:500:850").split(":");
		var url = "../../Common/" + (size[0] || "ItemSelectRelPartyCommon.asp") + "?" + buildQuery({
			orgID: field("hOrgID") ? field("hOrgID").value : "",
			sIType: field("hFromItemType") ? field("hFromItemType").value : "",
			Stock: "Y",
			hSelectMode: "S",
			Flag: "1",
			hDispButt: "",
			hDispItem: "",
			CallFrom: ""
		});
		openDialog(url, xmlIsland("ItemSelectData"), "dialogHeight:" + (size[1] || "500") + "px;dialogWidth:" + (size[2] || "850") + "px;Status:No", processSelectedMergeItem);
		return false;
	};

	window.GetXML = function () {
		var xhr;
		clearXml("OutData");
		xhr = syncGet("itmStatusXMLSelect.asp");
		if (trim(xhr.responseText)) {
			loadXmlIsland("OutData", xhr.responseText);
		}
		return false;
	};

	window.DisplayDetails = function () {
		var root = xmlRoot(xmlIsland("OutData"));
		var itemEntryNo = 0;
		rows = [];
		clearTable("tblData", 2);
		clearTable("tblToItemData", 1);
		removeChildren(tempRoot());
		allItems(root).forEach(function (item) {
			itemEntryNo += 1;
			if (field("hFromItemType")) {
				field("hFromItemType").value = getItemType(getAttr(item, "ICode"));
			}
			elementChildren(item, "LOCDET").forEach(function (loc) {
				renderSourceRow(item, loc, itemEntryNo);
			});
		});
		renderTargetRow();
		if (field("hItemRow")) {
			field("hItemRow").value = rows.length;
		}
		return false;
	};

	window.CalculateVal = function () {
		setText("spaMergedQty", previousMergeQuantity() + toNumber(textOf("spaMergeStock")));
		return false;
	};

	window.popClaDisplay = function () {
		return false;
	};

	window.clearXML = function () {
		clearXml("OutData");
		return false;
	};

	window.ClearTable = function () {
		clearTable("tblData", 2);
		return false;
	};

	window.CheckSubmit = function () {
		if (!field("hMergeItemCode") || trim(field("hMergeItemCode").value) === "") {
			alert("Select Item Merged With");
			return false;
		}
		buildIssueData();
		buildReceiptData();
		syncPost("XMLSave.asp?SessionFlag=true&Name=mrsIssueData", serializeXml(xmlIsland("IssueData")));
		syncPost("XMLSave.asp?SessionFlag=true&Name=ReceiptLotData", serializeXml(xmlIsland("IntReceipt")));
		form().action = "StkMergeInsert.asp";
		form().submit();
		return false;
	};

	window.GenReceiptXML = buildReceiptData;
	window.GenIssueXML = buildIssueData;
}(window, document));
