(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var rootDoc = null;
	var lineNo = "";
	var lotNo = "";
	var storeCode = "";
	var binCode = "";
	var orgId = "";
	var itemCode = "";
	var classCode = "";
	var itemEntryNo = "";

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
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function fieldAny() {
		for (var i = 0; i < arguments.length; i += 1) {
			if (field(arguments[i])) {
				return field(arguments[i]);
			}
		}
		return null;
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

	function formatNumber(value, decimals) {
		return toNumber(value).toFixed(decimals == null ? 3 : decimals);
	}

	function normalizeLot(value) {
		var lot = trim(value);
		return lot === "" || lot.toUpperCase() === "N/A" || lot.toUpperCase() === "NULL" ? "NULL" : lot;
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
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
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

	function createElement(name) {
		return (rootDoc || document).createElement(name);
	}

	function currentPick(createIfMissing) {
		var picks = elementChildren(root, "PICK");
		var normalizedLot = normalizeLot(lotNo);
		var i;
		for (i = 0; i < picks.length; i += 1) {
			if (normalizeLot(getAttr(picks[i], "LOTNO")) === normalizedLot &&
					getAttr(picks[i], "LOC") === trim(storeCode) &&
					getAttr(picks[i], "BIN") === trim(binCode)) {
				return picks[i];
			}
		}
		for (i = 0; i < picks.length; i += 1) {
			if (getAttr(picks[i], "LOC") === trim(storeCode) && getAttr(picks[i], "BIN") === trim(binCode)) {
				return picks[i];
			}
		}
		if (!createIfMissing || !root) {
			return null;
		}
		var pick = createElement("PICK");
		setAttr(pick, "ITEMENTRYNO", itemEntryNo);
		setAttr(pick, "LOC", storeCode);
		setAttr(pick, "BIN", binCode);
		setAttr(pick, "LOTNO", normalizedLot);
		setAttr(pick, "ISSQTY", "");
		root.appendChild(pick);
		return pick;
	}

	function allSelections() {
		var picks = elementChildren(root, "PICK");
		var result = [];
		var selections;
		for (var i = 0; i < picks.length; i += 1) {
			selections = elementChildren(picks[i], "Selection");
			for (var s = 0; s < selections.length; s += 1) {
				result.push(selections[s]);
			}
		}
		return result;
	}

	function currentSelections() {
		return elementChildren(currentPick(false), "Selection");
	}

	function selectedCurrentSelections() {
		var selections = currentSelections();
		var result = [];
		for (var i = 0; i < selections.length; i += 1) {
			if (getAttr(selections[i], "YesNo") === "Y") {
				result.push(selections[i]);
			}
		}
		return result;
	}

	function findSelectionBySerial(serialNo) {
		var selections = allSelections();
		for (var i = 0; i < selections.length; i += 1) {
			if (getAttr(selections[i], "SerialNo") === trim(serialNo)) {
				return selections[i];
			}
		}
		return null;
	}

	function removeSelectionBySerial(serialNo) {
		var selections = allSelections();
		var removed = false;
		for (var i = selections.length - 1; i >= 0; i -= 1) {
			if (getAttr(selections[i], "SerialNo") === trim(serialNo) && selections[i].parentNode) {
				selections[i].parentNode.removeChild(selections[i]);
				removed = true;
			}
		}
		return removed;
	}

	function rowCount() {
		return toNumber(field("hiCtr") && field("hiCtr").value);
	}

	function setSelectedText(total, count) {
		var span = byId("spaNoofPackSelected");
		if (span) {
			span.textContent = formatNumber(total, 3) + "[" + count + "]";
			span.innerText = span.textContent;
		}
		if (field("hRowSelect")) {
			field("hRowSelect").value = count;
		}
	}

	function refreshSelectedTotal() {
		var selections = selectedCurrentSelections();
		var total = 0;
		for (var i = 0; i < selections.length; i += 1) {
			total += toNumber(getAttr(selections[i], "Qty"));
		}
		setSelectedText(total, selections.length);
		return {
			total: total,
			count: selections.length
		};
	}

	function clearSelectedTable() {
		var table = byId("tblSerDet");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function appendCell(row, className, align, content) {
		var cell = row.insertCell(row.cells.length);
		cell.className = className;
		if (align) {
			cell.align = align;
		}
		if (typeof content === "string" || typeof content === "number") {
			cell.textContent = String(content);
		} else if (content) {
			cell.appendChild(content);
		}
		return cell;
	}

	function displaySelectedData() {
		var table = byId("tblSerDet");
		var selections;
		var row;
		var checkbox;
		clearSelectedTable();
		if (!table) {
			return;
		}
		selections = selectedCurrentSelections();
		for (var i = 0; i < selections.length; i += 1) {
			row = table.insertRow(table.rows.length);
			checkbox = document.createElement("input");
			checkbox.type = "checkbox";
			checkbox.name = "chkSerDet" + getAttr(selections[i], "SerialNo");
			checkbox.value = getAttr(selections[i], "SerialNo");
			checkbox.className = "FormElem";
			checkbox.checked = true;
			checkbox.style.textAlign = "center";
			checkbox.onclick = function () {
				window.RemoveNode(this);
			};
			appendCell(row, "ExcelInputCell", null, checkbox).width = "10";
			appendCell(row, "ExcelDisplayCell", "left", getAttr(selections[i], "PackNo"));
			appendCell(row, "ExcelDisplayCell", "right", getAttr(selections[i], "StockQty"));
			appendCell(row, "ExcelDisplayCell", "right", getAttr(selections[i], "Qty"));
		}
	}

	function removePagination(pick) {
		var pages = elementChildren(pick, "Pagination");
		for (var i = 0; i < pages.length; i += 1) {
			pick.removeChild(pages[i]);
		}
	}

	function paginationValue(status, page) {
		var searchBy = field("selSearchBy") && field("selSearchBy").options[field("selSearchBy").selectedIndex] ? field("selSearchBy").options[field("selSearchBy").selectedIndex].value : "";
		var searchFor = field("txtSearchFor") ? field("txtSearchFor").value : "";
		var searchType = field("selSearchType") ? field("selSearchType").value : "";
		var temp = field("hTemp") ? field("hTemp").value : "";
		var current = field("hCurrentPage") ? field("hCurrentPage").value : "";
		var rowCountField = fieldAny("hRowCount", "hRowcount");
		return temp + "|SearchBy=" + searchBy + "|SearchFor=" + searchFor + "|SearchType=" + searchType + "|hCurrentPage=" + current + "|hWho=" + status + "|hSubmit=" + (page || "") + "|hRowCount=" + (rowCountField ? rowCountField.value : "");
	}

	function xmlSaveMod() {
		var mod = field("hXMLSaveMod");
		return mod && trim(mod.value) ? trim(mod.value) : "IssuePickPack";
	}

	function addPagination(status, page) {
		var pick = currentPick(true);
		var pagination = createElement("Pagination");
		removePagination(pick);
		setAttr(pagination, "Details", paginationValue(status, page));
		pick.appendChild(pagination);
	}

	function setReturnValue() {
		if (!root) {
			root = xmlRoot(objTemp);
		}
		if (!root) {
			return;
		}
		window.returnValue = root;
		window.returnvalue = root;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	window.checkNumbers = function (value) {
		return /^[0-9.]+$/.test(String(value || ""));
	};

	window.fnInit = function (arg) {
		var parts = String(arg || "").split(":");
		var total = 0;
		var selections;

		objTemp = window.dialogArguments;
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		if (!rootDoc && root) {
			rootDoc = root.ownerDocument;
		}

		lineNo = parts[1] || "";
		lotNo = normalizeLot(parts[2] || "");
		storeCode = parts[3] || "";
		binCode = parts[4] || "";
		classCode = parts[6] || "";
		itemCode = parts[7] || "";
		orgId = parts[8] || "";
		itemEntryNo = parts[10] || "";

		if (field("hRowSelect")) {
			field("hRowSelect").value = getAttr(root, "NoofPack") || 0;
		}

		selections = currentSelections();
		for (var i = selections.length - 1; i >= 0; i -= 1) {
			if (getAttr(selections[i], "YesNo") === "N") {
				selections[i].parentNode.removeChild(selections[i]);
			}
		}
		selections = selectedCurrentSelections();
		for (var s = 0; s < selections.length; s += 1) {
			total += toNumber(getAttr(selections[s], "Qty"));
		}
		setSelectedText(total, selections.length);
		displaySelectedData();
	};

	window.CheckSer = function () {
		var count = rowCount();
		var removed = false;
		var checkbox;
		var qty;
		var stock;
		var serial;
		for (var i = 1; i <= count; i += 1) {
			checkbox = field("chkSer" + i);
			qty = field("txtQty" + i);
			stock = field("txtStQty" + i);
			if (!checkbox || !qty) {
				continue;
			}
			if (checkbox.checked) {
				if (qty.value === "0" || trim(qty.value) === "") {
					qty.value = stock ? stock.value : qty.value;
				}
			} else {
				qty.value = "0";
				serial = field("hSerial" + i);
				if (serial && removeSelectionBySerial(serial.value)) {
					removed = true;
				}
			}
		}
		if (removed) {
			displaySelectedData();
			refreshSelectedTotal();
		}
	};

	window.btnAddToList_Click = function () {
		var count = rowCount();
		var pick = currentPick(true);
		var selectedCount = 0;
		var checkbox;
		var selection;
		var serial;
		var stockQty;
		var qty;
		var packNo;
		var rowLot;

		if (!count) {
			alert("Select the Pack Numbers");
			return;
		}

		for (var i = 1; i <= count; i += 1) {
			checkbox = field("chkSer" + i);
			if (!checkbox || !checkbox.checked) {
				continue;
			}
			serial = field("hSerial" + i);
			stockQty = field("txtStQty" + i);
			qty = field("txtQty" + i);
			packNo = field("txtSerial" + i);
			rowLot = field("hLotNoZ" + i);
			if (!serial) {
				continue;
			}
			removeSelectionBySerial(serial.value);
			selection = createElement("Selection");
			setAttr(selection, "SerialNo", serial.value);
			setAttr(selection, "StockQty", stockQty ? stockQty.value : "");
			setAttr(selection, "Qty", qty ? qty.value : "");
			setAttr(selection, "YesNo", "Y");
			setAttr(selection, "PackNo", packNo ? packNo.value : serial.value);
			setAttr(selection, "LotNo", rowLot ? rowLot.value : normalizeLot(lotNo));
			pick.appendChild(selection);
			selectedCount += 1;
		}

		if (!selectedCount || refreshSelectedTotal().total <= 0) {
			alert("Select the Pack Numbers");
			return;
		}
		displaySelectedData();
	};

	window.TotQtyDisplay = refreshSelectedTotal;
	window.ClearTable = clearSelectedTable;
	window.DisplayData = displaySelectedData;

	window.RemoveNode = function (control) {
		var serial = control && control.value;
		var count = rowCount();
		for (var i = 1; i <= count; i += 1) {
			if (field("hSerial" + i) && trim(field("hSerial" + i).value) === trim(serial)) {
				if (field("chkSer" + i)) {
					field("chkSer" + i).checked = false;
				}
				if (field("txtQty" + i)) {
					field("txtQty" + i).value = "0";
				}
			}
		}
		removeSelectionBySerial(serial);
		displaySelectedData();
		refreshSelectedTotal();
	};

	window.btnRemove_Click = function () {
		var select = field("selFrombox");
		var options = select ? select.options : [];
		for (var i = 0; i < options.length; i += 1) {
			if (options[i].selected) {
				var selection = findSelectionBySerial(options[i].value);
				if (selection) {
					setAttr(selection, "YesNo", "N");
				}
			}
		}
		refreshSelectedTotal();
	};

	window.btnAdd_Click = function () {
		var select = field("selTobox");
		var options = select ? select.options : [];
		for (var i = 0; i < options.length; i += 1) {
			if (options[i].selected) {
				var selection = findSelectionBySerial(options[i].value);
				if (selection) {
					setAttr(selection, "YesNo", "Y");
				}
			}
		}
		refreshSelectedTotal();
	};

	window.CheckSubmit = function () {
		var pick = currentPick(true);
		var selections = selectedCurrentSelections();
		var total = 0;
		var serialHeader;
		var serialDetail;
		var oldHeaders;
		var subCon = field("hSUBC") ? field("hSUBC").value : "";

		if (String(subCon).toLowerCase() === "no") {
			for (var i = 0; i < selections.length; i += 1) {
				total += toNumber(getAttr(selections[i], "Qty"));
			}
		}

		setAttr(root, "NoofPack", selections.length);
		oldHeaders = elementChildren(pick, "SERIALHEADER");
		for (var h = 0; h < oldHeaders.length; h += 1) {
			pick.removeChild(oldHeaders[h]);
		}

		serialHeader = createElement("SERIALHEADER");
		for (var s = 0; s < selections.length; s += 1) {
			serialDetail = createElement("SERIALDETAILS");
			setAttr(serialDetail, "SERIALNO", getAttr(selections[s], "SerialNo") || "NULL");
			setAttr(serialDetail, "QTY", getAttr(selections[s], "Qty"));
			setAttr(serialDetail, "LOTNO", getAttr(selections[s], "LotNo"));
			serialHeader.appendChild(serialDetail);
		}
		pick.appendChild(serialHeader);
		setAttr(pick, "ISSQTY", total);
		addPagination("DONE```YES", "");
		closeWithReturn();
	};

	window.btnClose_Click = function () {
		addPagination("DONE```NO", "");
		closeWithReturn();
	};

	window.NextSelection = function (page) {
		addPagination("NEXT", page);
		closeWithReturn();
	};

	window.CheckSerialNo = function () {
		var count = rowCount();
		var selection;
		var serial;
		for (var i = 1; i <= count; i += 1) {
			serial = field("hSerial" + i);
			if (!serial) {
				continue;
			}
			selection = findSelectionBySerial(serial.value);
			if (selection) {
				if (field("chkSer" + i)) {
					field("chkSer" + i).checked = true;
				}
				if (field("txtQty" + i)) {
					field("txtQty" + i).value = getAttr(selection, "Qty");
				}
			}
		}
	};

	window.Save = function () {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSave.asp?Name=Det&Mod=" + encodeURIComponent(xmlSaveMod()), false);
		xhr.send(rootDoc || objTemp);
	};

	window.CheckAll = function () {
		var count = rowCount();
		var checked = field("chkAll") && field("chkAll").checked;
		for (var i = 1; i <= count; i += 1) {
			if (field("chkSer" + i)) {
				field("chkSer" + i).checked = checked;
			}
		}
		window.CheckSer();
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
