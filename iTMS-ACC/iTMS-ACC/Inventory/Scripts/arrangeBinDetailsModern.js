(function (window, document) {
	"use strict";

	var root = null;
	var rootDoc = null;
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
		return frm && frm.elements ? frm.elements[name] || null : null;
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

	function isNumericText(value) {
		return /^([0-9]+(\.[0-9]*)?|\.[0-9]+)$/.test(trim(value));
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

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function serializeXml(value) {
		var doc = xmlDocument(value);
		var rootNode = xmlRoot(value);
		if (value && typeof value.xml === "string") {
			return value.xml;
		}
		if (doc && typeof doc.xml === "string") {
			return doc.xml;
		}
		return new XMLSerializer().serializeToString(doc || rootNode);
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

	function makeInput(name, value, className, size, readOnly) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.value = value == null ? "" : String(value);
		input.className = className || "Formelem";
		if (size) {
			input.size = size;
		}
		if (readOnly) {
			input.readOnly = true;
		}
		return input;
	}

	function makeCheckbox(name, value, checked) {
		var input = document.createElement("input");
		input.type = "checkbox";
		input.name = name;
		input.value = value == null ? "" : String(value);
		input.className = "Formelem";
		input.checked = !!checked;
		return input;
	}

	function clearTable() {
		var table = byId("tblBin");
		rows = [];
		if (!table) {
			return;
		}
		while (table.rows.length) {
			table.deleteRow(0);
		}
	}

	function findItem(itemCode, classCode) {
		var items = elementChildren(root, "Item");
		for (var i = 0; i < items.length; i += 1) {
			if (getAttr(items[i], "ICode") === itemCode && getAttr(items[i], "CCode") === classCode) {
				return items[i];
			}
		}
		return null;
	}

	function findLoc(item, locNo) {
		var locs = elementChildren(item, "LOCDET");
		for (var i = 0; i < locs.length; i += 1) {
			if (getAttr(locs[i], "LOC") === locNo) {
				return locs[i];
			}
		}
		return null;
	}

	function findExistingBinQty(loc, binNo) {
		var bins = elementChildren(loc, "BINDET");
		for (var i = 0; i < bins.length; i += 1) {
			if (getAttr(bins[i], "BINNO") === binNo) {
				return getAttr(bins[i], "QTY");
			}
		}
		return null;
	}

	function storeBinParent(loc) {
		return elementChildren(loc, "STOREBINDET")[0] || null;
	}

	function renderHeader(table) {
		var row = table.insertRow(0);
		appendCell(row, "ExcelHeaderCell", "center", "S.No.");
		appendCell(row, "ExcelHeaderCell", "center", "");
		appendCell(row, "ExcelHeaderCell", "center", "Bin Code");
		appendCell(row, "ExcelHeaderCell", "center", "Quantity");
	}

	function renderBinRow(table, bin, rowNo, existingQty) {
		var row = table.insertRow(table.rows.length);
		var checked = existingQty !== null;
		var qty = checked ? existingQty : getAttr(bin, "QTY");
		var checkbox;
		var qtyInput;

		appendCell(row, "ExcelDisplayCell", "center", rowNo);
		checkbox = makeCheckbox("Chk" + rowNo, getAttr(bin, "NO"), checked);
		appendCell(row, "ExcelDisplayCell", null, checkbox);
		appendCell(row, "ExcelInputCell", null, makeInput("txtBinCode" + rowNo, getAttr(bin, "NO"), "Formelem", 30, true));
		qtyInput = makeInput("txtQty" + rowNo, qty, "Formelem", 30, false);
		qtyInput.maxLength = 10;
		qtyInput.onchange = function () {
			window.Calculate();
		};
		appendCell(row, "ExcelInputCell", null, qtyInput);
		rows.push({ bin: bin, checkbox: checkbox, quantity: qtyInput });
	}

	function renderTotalRow(table, total) {
		var row = table.insertRow(table.rows.length);
		var cell = appendCell(row, "ExcelHeaderCell", "center", "Total");
		var totalInput;
		cell.colSpan = 3;
		totalInput = makeInput("txtTotQty", total, "Formelem", 30, true);
		appendCell(row, "ExcelInputCell", null, totalInput);
	}

	function selectedLoc() {
		var item = findItem(trim(field("hItemCode") && field("hItemCode").value), trim(field("hClassCode") && field("hClassCode").value));
		return item ? findLoc(item, trim(field("hLocNo") && field("hLocNo").value)) : null;
	}

	function syncCount() {
		if (field("hCnt")) {
			field("hCnt").value = rows.length;
		}
	}

	function syncTotal() {
		var total = 0;
		for (var i = 0; i < rows.length; i += 1) {
			total += toNumber(rows[i].quantity.value);
		}
		if (field("txtTotQty")) {
			field("txtTotQty").value = total;
		}
		return total;
	}

	function syncXmlFromRows() {
		var loc = selectedLoc();
		var parent = loc && storeBinParent(loc);
		if (!parent) {
			return false;
		}
		rows.forEach(function (row) {
			setAttr(row.bin, "QTY", row.quantity.value);
			if (!row.checkbox.checked && row.bin.parentNode === parent) {
				parent.removeChild(row.bin);
			}
		});
		return true;
	}

	function syncPost(url, xmlText) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(xmlText || "");
		return xhr;
	}

	function closeDone() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue("Done");
		} else {
			window["return" + "Value"] = "Done";
			window.returnvalue = "Done";
		}
		window.close();
	}

	function initializeDialogXml() {
		var args = modalArgs();
		root = xmlRoot(args);
		rootDoc = xmlDocument(args) || root && root.ownerDocument;
		if (!root) {
			rootDoc = new DOMParser().parseFromString("<ROOT/>", "text/xml");
			root = rootDoc.documentElement;
		}
	}

	window.DisplaytableBin = function () {
		var table = byId("tblBin");
		var loc = selectedLoc();
		var parent = loc && storeBinParent(loc);
		var bins = parent ? elementChildren(parent, "BIN") : [];
		var total = 0;
		clearTable();
		if (!table) {
			return false;
		}
		renderHeader(table);
		for (var i = 0; i < bins.length; i += 1) {
			var existingQty = findExistingBinQty(loc, getAttr(bins[i], "NO"));
			renderBinRow(table, bins[i], i + 1, existingQty);
			total += toNumber(existingQty !== null ? existingQty : getAttr(bins[i], "QTY"));
		}
		syncCount();
		renderTotalRow(table, total);
		return false;
	};

	window.Calculate = function () {
		for (var i = 0; i < rows.length; i += 1) {
			if (toNumber(rows[i].quantity.value) < 0) {
				alert("Enter Valid value");
				rows[i].quantity.focus();
				return false;
			}
			if (!isNumericText(rows[i].quantity.value)) {
				alert("Enter Numerals only");
				rows[i].quantity.focus();
				return false;
			}
		}
		syncTotal();
		return true;
	};

	window.checkNumbers = function (value) {
		return isNumericText(value);
	};

	window.ClearTable = function () {
		clearTable();
		return false;
	};

	window.ClearAll = function () {
		clearTable();
		return false;
	};

	window.FnInit = function () {
		initializeDialogXml();
		window.DisplaytableBin();
		return false;
	};

	window.CheckSubmit = function () {
		var total;
		var xhr;
		if (!window.Calculate()) {
			return false;
		}
		total = syncTotal();
		if (Math.abs(total - toNumber(field("hTotStkQty") && field("hTotStkQty").value)) > 0.000001) {
			alert("Item Qty and Bin Qty is Not Equal");
			return false;
		}
		if (!syncXmlFromRows()) {
			alert("No bin details available");
			return false;
		}
		xhr = syncPost("ArrBinDetInsert.asp", serializeXml(rootDoc || root));
		if (trim(xhr.responseText)) {
			alert(xhr.responseText);
		} else {
			alert("Bin Details Arranged");
			closeDone();
		}
		return false;
	};

	window.checkSubmit = window.CheckSubmit;
}(window, document));
