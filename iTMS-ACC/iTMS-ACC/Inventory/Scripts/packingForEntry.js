(function (window, document) {
	"use strict";

	var serialCounter = 0;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		var item;
		var wanted;
		var i;
		if (!frm || !frm.elements) {
			return null;
		}
		item = frm.elements[name];
		if (item) {
			return item;
		}
		wanted = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === wanted) {
				return frm.elements[i];
			}
		}
		return null;
	}

	function valueOf(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function asNumber(value) {
		var number = parseFloat(value);
		return isNaN(number) ? 0 : number;
	}

	function checkNumbers(value) {
		return /^[0-9]+(?:\.[0-9]+)?$/.test(String(value || ""));
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.xmlRoot) {
			return window.ITMSModalReturnCompat.xmlRoot(object);
		}
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		var root = xmlRoot(object);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
	}

	function childElements(node, name) {
		var result = [];
		var wanted = name ? String(name).toUpperCase() : "";
		var child;
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			child = node.childNodes[i];
			if (child.nodeType === 1 && (!wanted || String(child.nodeName || "").toUpperCase() === wanted)) {
				result.push(child);
			}
		}
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function serializeXml(rootOrDoc) {
		var doc = rootOrDoc && rootOrDoc.nodeType === 9 ? rootOrDoc : rootOrDoc && rootOrDoc.ownerDocument;
		return new XMLSerializer().serializeToString(doc || rootOrDoc);
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback || function () {});
		}
		window.open(url, "_blank");
		return null;
	}

	function show(id, visible) {
		var item = document.getElementById(id);
		if (item) {
			item.style.display = visible ? "block" : "none";
		}
	}

	function setText(id, value) {
		var item = document.getElementById(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function Save() {
		var root = xmlRoot("LotData");
		alert(root ? serializeXml(root) : "");
		return false;
	}

	function ClearTable() {
		var table = document.getElementById("tabNoOfLevel");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function AddSubLevel() {
		var count = Number(valueOf("txtNoOfSubLevel")) || 0;
		var table = document.getElementById("tabNoOfLevel");
		var row;
		var cell;
		var input;
		ClearTable();
		if (!table) {
			return false;
		}
		for (var i = 1; i <= count; i += 1) {
			row = table.insertRow(-1);
			cell = row.insertCell(-1);
			cell.textContent = i;
			cell.className = "ExcelHeaderCell";
			cell.align = "center";
			cell = row.insertCell(-1);
			cell.className = "ExcelDisplayCell";
			input = document.createElement("input");
			input.type = "text";
			input.name = "txtLevelLable" + i;
			input.value = "";
			input.className = "FormElem";
			cell.appendChild(input);
		}
		return false;
	}

	function ViewInvoice() {
		var value = valueOf("hApplnTransNo");
		if (value && value !== "0") {
			openDialog("../../Purchase/Transaction/RepPurInvoiceDetailspopup.asp?iInvNo=" + encodeURIComponent(value), "", "dialogHeight:470px;dialogWidth:870px;center:Yes;help:No;resizable:No;status:No");
		}
		return false;
	}

	function ShowSubLevelQty() {
		openDialog("SubLevelQtyPopUp.asp", xmlObject("SubLevelQty"), "dialogHeight:230px;dialogWidth:370px;center:Yes;help:No;resizable:No;status:No");
		return false;
	}

	function CheckLotSerial(select) {
		var parts = String(select && select.value || "").split(":");
		var numbering = parts[1] || "";
		var packShortName = parts[2] || "";
		var mode;
		if (!select || select.value === "S") {
			return false;
		}
		if (numbering === "L") {
			mode = parts[3] || "";
			setText("Data3", "Lot");
			setText("LastUsedData", valueOf("hLastUsedLotNo"));
			field("selLotNo").disabled = mode !== "A";
			show("DivAuto", mode === "A");
			show("DivManual", mode !== "A");
			show("DivLastUsed", mode !== "A");
			show("DivMain", true);
		} else if (numbering === "S") {
			mode = parts[4] || "";
			setText("Data3", "Serial");
			setText("LastUsedData", "123");
			field("selLotNo").disabled = false;
			show("DivMain", mode !== "A");
			show("DivAuto", false);
			show("DivManual", mode !== "A");
			show("DivLastUsed", true);
		}
		setText("Data1", packShortName);
		setText("Data2", packShortName);
		return false;
	}

	function packNodeFor(root, packCode) {
		var found = null;
		childElements(root, "PackDetails").some(function (node) {
			if (trim(attr(node, "PackingCode")) === trim(packCode)) {
				found = node;
				return true;
			}
			return false;
		});
		return found;
	}

	function removeChildrenByName(node, name) {
		childElements(node, name).forEach(function (child) {
			node.removeChild(child);
		});
	}

	function selectedPackParts() {
		return String(valueOf("selPackType")).split(":");
	}

	function selectedLotNo(parts) {
		var numbering = parts[1] || "";
		var lotMode = parts[3] || "";
		var serialMode = parts[4] || "";
		if (numbering === "L") {
			if (lotMode === "A") {
				if (valueOf("selLotNo") === "S") {
					alert("Select Lot Number");
					field("selLotNo").focus();
					return null;
				}
				return valueOf("selLotNo");
			}
			if (!valueOf("txtLotNo")) {
				alert("Enter Lot Number");
				field("txtLotNo").focus();
				return null;
			}
			return valueOf("txtLotNo");
		}
		if (numbering === "S") {
			if (serialMode === "A") {
				return "123";
			}
			if (!valueOf("txtLotNo")) {
				alert("Enter Lot Number");
				field("txtLotNo").focus();
				return null;
			}
			return valueOf("txtLotNo");
		}
		return valueOf("selLotNo") || valueOf("txtLotNo") || "";
	}

	function AddData() {
		var parts = selectedPackParts();
		var packCode = parts[0] || "";
		var numbering = parts[1] || "";
		var packCount = Number(valueOf("txtNoOfPack")) || 0;
		var grossPerPack = asNumber(valueOf("txtuniSerQty"));
		var tare = valueOf("txtTareQty") ? asNumber(valueOf("txtTareQty")) : 0;
		var noOfBag = valueOf("txtSubLevelQty");
		var indKgs = valueOf("txtSubLevelwt");
		var lotNo;
		var root;
		var doc;
		var packNode;
		var qtyNode;
		var nettQty;
		if (valueOf("selPackType") === "S") {
			alert("select Pack Type");
			field("selPackType").focus();
			return false;
		}
		if (!valueOf("txtNoOfPack")) {
			alert("Enter No Of Pack");
			field("txtNoOfPack").focus();
			return false;
		}
		if (!checkNumbers(valueOf("txtNoOfPack"))) {
			alert("Enter numerals");
			field("txtNoOfPack").focus();
			return false;
		}
		if (!valueOf("txtuniSerQty")) {
			alert("Enter Gross Qty");
			field("txtuniSerQty").focus();
			return false;
		}
		if (!checkNumbers(valueOf("txtuniSerQty"))) {
			alert("Enter numerals");
			field("txtuniSerQty").focus();
			return false;
		}
		lotNo = selectedLotNo(parts);
		if (lotNo == null) {
			return false;
		}
		nettQty = grossPerPack - tare;
		root = xmlRoot("LotData");
		doc = xmlDocument("LotData");
		packNode = packNodeFor(root, packCode);
		if (!packNode) {
			packNode = doc.createElement("PackDetails");
			packNode.setAttribute("PackingCode", packCode);
			root.appendChild(packNode);
		}
		removeChildrenByName(packNode, "QtyDetails");
		packNode.setAttribute("RecNumbering", numbering);
		packNode.setAttribute("NoOfPack", packCount);
		packNode.setAttribute("IndKgs", indKgs);
		packNode.setAttribute("LotNo", lotNo);
		packNode.setAttribute("TotalQty", packCount * grossPerPack);
		for (var i = 1; i <= packCount; i += 1) {
			serialCounter += 1;
			qtyNode = doc.createElement("QtyDetails");
			qtyNode.setAttribute("Cnt", serialCounter);
			qtyNode.setAttribute("GrossQty", grossPerPack);
			qtyNode.setAttribute("NettQty", nettQty);
			qtyNode.setAttribute("TareQty", tare || "");
			qtyNode.setAttribute("NoOfBagorCone", noOfBag);
			qtyNode.setAttribute("PartySerNo", i);
			qtyNode.setAttribute("PartySerQty", "");
			packNode.appendChild(qtyNode);
		}
		DispalyTable();
		return false;
	}

	function clearTable() {
		var table = document.getElementById("tblLotDetail");
		if (!table) {
			return;
		}
		while (table.rows.length > 2) {
			table.deleteRow(2);
		}
	}

	function addInput(cell, name, value, size, readOnly, onChange) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.size = size || 11;
		input.value = value == null ? "" : String(value);
		input.className = readOnly ? "FormElemRead" : "FormElem";
		input.readOnly = !!readOnly;
		if (onChange) {
			input.onchange = onChange;
		}
		cell.appendChild(input);
		return input;
	}

	function addCell(row, text, className, align) {
		var cell = row.insertCell(-1);
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		if (text != null) {
			cell.textContent = String(text);
		}
		return cell;
	}

	function populatePackSelect(select, selectedCode) {
		var root = xmlRoot("PACKFORMDATA");
		var option;
		select.length = 0;
		childElements(root, "PACK").forEach(function (node) {
			option = document.createElement("option");
			option.value = attr(node, "VALUE");
			option.text = attr(node, "NAME");
			option.selected = trim(selectedCode) === trim(option.value);
			select.add(option);
		});
		if (!select.length) {
			option = document.createElement("option");
			option.value = "0";
			option.text = " -NA-";
			select.add(option);
		}
	}

	function addPackSelect(cell, name, selectedCode) {
		var select = document.createElement("select");
		select.size = 1;
		select.name = name;
		select.className = "FormElem";
		populatePackSelect(select, selectedCode);
		cell.appendChild(select);
	}

	function recalcTotals() {
		var totalGross = 0;
		var totalNett = 0;
		var totalBag = 0;
		var count = Number(valueOf("hCnt")) || 0;
		for (var i = 1; i <= count; i += 1) {
			totalGross += asNumber(valueOf("txtGrossQtyZ" + i));
			totalNett += asNumber(valueOf("txtNettQtyZ" + i));
			totalBag += asNumber(valueOf("txtNoOfBagZ" + i));
		}
		setValue("txtTotGrossQty", Math.round(totalGross * 100) / 100);
		setValue("txtTotNettQty", Math.round(totalNett * 100) / 100);
		setValue("txtTotNoOfBag", totalBag);
	}

	function calcGrossTareNettWt(packCode, lotNo, count, type) {
		var root = xmlRoot("LotData");
		var packNode = packNodeFor(root, packCode);
		var firstQty = childElements(packNode, "QtyDetails")[0];
		var tare = asNumber(attr(firstQty, "TareQty"));
		var gross;
		var nett;
		if (type === "G") {
			gross = valueOf("txtGrossQtyZ" + count);
			if (!checkNumbers(gross)) {
				alert("Enter Numeric Gross Weight");
				field("txtGrossQtyZ" + count).focus();
				return false;
			}
			setValue("txtNettQtyZ" + count, asNumber(gross) - tare);
		} else if (type === "N") {
			nett = valueOf("txtNettQtyZ" + count);
			if (!checkNumbers(nett)) {
				alert("Enter Numeric Nett Weight");
				field("txtNettQtyZ" + count).focus();
				return false;
			}
			setValue("txtGrossQtyZ" + count, asNumber(nett) + tare);
		}
		recalcTotals();
		return false;
	}

	function addDetailRows(table, packNode) {
		var totalGross = 0;
		var totalNett = 0;
		var totalBag = 0;
		var packCode = attr(packNode, "PackingCode");
		var lotNo = attr(packNode, "LotNo");
		var header = table.insertRow(-1);
		var cell = header.insertCell(-1);
		cell.colSpan = 9;
		cell.className = "ExcelDisplayCell";
		addInput(cell, "txtLotNoZ" + lotNo, lotNo + "-" + attr(packNode, "TotalQty") + "[]", 110, true);
		childElements(packNode, "QtyDetails").forEach(function (node) {
			var count = Number(attr(node, "Cnt")) || 0;
			var row = table.insertRow(-1);
			var checkbox;
			var selectCell;
			serialCounter = Math.max(serialCounter, count);
			totalGross += asNumber(attr(node, "GrossQty"));
			totalNett += asNumber(attr(node, "NettQty"));
			totalBag += asNumber(attr(node, "NoOfBagorCone"));
			cell = addCell(row, "", "ExcelDisplayCell", "left");
			addInput(cell, "txtMillSlNoZ" + count, count, 4, true);
			cell = addCell(row, "", "ExcelDisplayCell", "center");
			checkbox = document.createElement("input");
			checkbox.type = "checkbox";
			checkbox.name = "chkBoxZ" + count;
			checkbox.value = packCode + ":" + count;
			checkbox.className = "FormElemRead";
			cell.appendChild(checkbox);
			selectCell = addCell(row, "", "ExcelDisplayCell", "center");
			addPackSelect(selectCell, "cmbPackTypeZ" + count, packCode);
			cell = addCell(row, "", "ExcelInputCell", "right");
			addInput(cell, "txtGrossQtyZ" + count, attr(node, "GrossQty"), 11, false, function () {
				calcGrossTareNettWt(packCode, lotNo, count, "G");
			});
			cell = addCell(row, "", "ExcelInputCell", "right");
			addInput(cell, "txtNettQtyZ" + count, attr(node, "NettQty"), 11, false, function () {
				calcGrossTareNettWt(packCode, lotNo, count, "N");
			});
			cell = addCell(row, "", "ExcelInputCell", "right");
			addInput(cell, "txtNoOfBagZ" + count, attr(node, "NoOfBagorCone"), 11, false, recalcTotals);
			cell = addCell(row, "", "ExcelInputCell", "left");
			addInput(cell, "txtPartySerialZ" + count, attr(node, "PartySerNo"), 11, false);
			cell = addCell(row, "", "ExcelInputCell", "left");
			addInput(cell, "txtPartyQtyZ" + count, attr(node, "PartySerQty"), 11, false);
			setValue("hCnt", count);
		});
		addTotalRow(table, totalGross, totalNett, totalBag);
	}

	function addTotalRow(table, gross, nett, bag) {
		var row = table.insertRow(-1);
		var cell = addCell(row, "Total", "ExcelDisplayCell", "right");
		cell.colSpan = 3;
		cell = addCell(row, "", "ExcelInputCell", "right");
		addInput(cell, "txtTotGrossQty", gross, 11, true);
		cell = addCell(row, "", "ExcelInputCell", "right");
		addInput(cell, "txtTotNettQty", nett, 11, true);
		cell = addCell(row, "", "ExcelInputCell", "right");
		addInput(cell, "txtTotNoOfBag", bag, 11, true);
		cell = addCell(row, "", "ExcelDisplayCell", "left");
		cell.colSpan = 3;
	}

	function DispalyTable() {
		var root = xmlRoot("LotData");
		var table = document.getElementById("tblLotDetail");
		serialCounter = 0;
		clearTable();
		childElements(root, "PackDetails").forEach(function (node) {
			addDetailRows(table, node);
		});
	}

	function DeleteData() {
		var root = xmlRoot("LotData");
		var count = Number(valueOf("hCnt")) || 0;
		var removeByPack = {};
		var item;
		var parts;
		var nextCount = 1;
		for (var i = 1; i <= count; i += 1) {
			item = field("chkBoxZ" + i);
			if (item && item.checked) {
				parts = String(item.value || "").split(":");
				removeByPack[parts[0]] = removeByPack[parts[0]] || {};
				removeByPack[parts[0]][parts[1]] = true;
			}
		}
		childElements(root, "PackDetails").forEach(function (packNode) {
			var packCode = attr(packNode, "PackingCode");
			childElements(packNode, "QtyDetails").forEach(function (qtyNode) {
				if (removeByPack[packCode] && removeByPack[packCode][attr(qtyNode, "Cnt")]) {
					packNode.removeChild(qtyNode);
				}
			});
		});
		childElements(root, "PackDetails").forEach(function (packNode) {
			var total = 0;
			childElements(packNode, "QtyDetails").forEach(function (qtyNode, index) {
				qtyNode.setAttribute("Cnt", nextCount);
				nextCount += 1;
				total += asNumber(attr(qtyNode, "GrossQty"));
			});
			packNode.setAttribute("TotalQty", total);
		});
		DispalyTable();
		return false;
	}

	window.Save = Save;
	window.AddSubLevel = AddSubLevel;
	window.ClearTable = ClearTable;
	window.ViewInvoice = ViewInvoice;
	window.ShowSubLevelQty = ShowSubLevelQty;
	window.CheckLotSerial = CheckLotSerial;
	window.AddData = AddData;
	window.clearTable = clearTable;
	window.DispalyTable = DispalyTable;
	window.calcGrossTareNettWt = calcGrossTareNettWt;
	window.DeleteData = DeleteData;
	window.checkNumbers = checkNumbers;
}(window, document));
