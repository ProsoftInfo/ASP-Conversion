(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function upper(value) {
		return trim(value).toUpperCase();
	}

	function numberValue(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value, decimals) {
		return numberValue(value).toFixed(decimals == null ? 2 : decimals);
	}

	function round(value, decimals) {
		var factor = Math.pow(10, decimals || 0);
		return Math.round(numberValue(value) * factor) / factor;
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		if (!frm) {
			return null;
		}
		return frm.elements[name] || frm.elements[String(name).toLowerCase()] || frm.elements[String(name).toUpperCase()] || null;
	}

	function setField(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function setText(id, value) {
		var el = document.getElementById(id) || window[id];
		if (el) {
			el.textContent = value == null ? "" : String(value);
		}
	}

	function textValue(id) {
		var el = document.getElementById(id) || window[id];
		return trim(el && (el.textContent || el.innerText || ""));
	}

	function xmlDoc(wrapper) {
		if (!wrapper) {
			return null;
		}
		if (wrapper.XMLDocument) {
			return wrapper.XMLDocument;
		}
		if (wrapper._doc) {
			return wrapper._doc;
		}
		if (wrapper.nodeType === 9) {
			return wrapper;
		}
		if (wrapper.ownerDocument) {
			return wrapper.ownerDocument;
		}
		return null;
	}

	function xmlRoot(wrapper) {
		if (!wrapper) {
			return null;
		}
		if (wrapper.documentElement) {
			return wrapper.documentElement;
		}
		if (wrapper.nodeType === 1) {
			return wrapper;
		}
		var doc = xmlDoc(wrapper);
		return doc && doc.documentElement;
	}

	function invoiceDoc() {
		return xmlDoc(window.InvoiceDet);
	}

	function invoiceRoot() {
		return xmlRoot(window.InvoiceDet);
	}

	function createElement(name) {
		return window.InvoiceDet && typeof window.InvoiceDet.createElement === "function" ?
			window.InvoiceDet.createElement(name) :
			invoiceDoc().createElement(name);
	}

	function selectNodes(context, expression) {
		return context && typeof context.selectNodes === "function" ? context.selectNodes(expression) : [];
	}

	function xpathLiteral(value) {
		var text = String(value == null ? "" : value);
		return text.indexOf("'") === -1 ? "'" + text + "'" : '"' + text.replace(/"/g, "") + '"';
	}

	function childElements(node) {
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1;
		});
	}

	function attr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function setAttr(node, name, value) {
		if (node) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function firstNode(nodes) {
		return nodes && nodes.length ? nodes.item(0) : null;
	}

	function headerNode() {
		return firstNode(selectNodes(invoiceRoot(), "//InvoiceHeader/Header"));
	}

	function itemNodes(purchaseType) {
		var root = invoiceRoot();
		if (purchaseType == null || purchaseType === "") {
			return selectNodes(root, "//ItemDetails/Item");
		}
		return selectNodes(root, "//ItemDetails/Item[@PurchaseType=" + xpathLiteral(purchaseType) + "]");
	}

	function taxRoots(purchaseType) {
		var root = invoiceRoot();
		if (purchaseType == null || purchaseType === "") {
			return selectNodes(root, "//TaxDetails");
		}
		return selectNodes(root, "//TaxDetails[@PurchaseType=" + xpathLiteral(purchaseType) + "]");
	}

	function firstTaxRoot(purchaseType) {
		return firstNode(taxRoots(purchaseType));
	}

	function taxNodeByCodes(taxRoot, catCode, taxCode) {
		var nodes = childElements(taxRoot);
		for (var index = 0; index < nodes.length; index += 1) {
			if (attr(nodes[index], "CatCode") === trim(catCode) && attr(nodes[index], "TaxCode") === trim(taxCode)) {
				return nodes[index];
			}
		}
		return null;
	}

	function xmlText(value) {
		var target = value && (value.XMLDocument || value._doc || value);
		if (!target) {
			return "";
		}
		if (typeof target.xml === "string") {
			return target.xml;
		}
		return new XMLSerializer().serializeToString(target);
	}

	function importForInvoice(node) {
		var doc = invoiceDoc();
		return node.ownerDocument === doc ? node : doc.importNode(node, true);
	}

	function loadInvoiceXml(text) {
		if (window.InvoiceDet && typeof window.InvoiceDet.loadXML === "function") {
			window.InvoiceDet.loadXML(text);
		}
	}

	function openDialog(url, args, features, callback) {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function saveInvoiceXml() {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSavePur.asp?Mod=PUR&Name=AmdNewInvItemValue", false);
		xhr.setRequestHeader("Content-Type", "text/xml; charset=UTF-8");
		xhr.send(xmlText(window.InvoiceDet));
		return xhr;
	}

	function removeChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function selectedValue(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function selectedText(name) {
		var item = field(name);
		return item && item.options && item.selectedIndex >= 0 ? item.options[item.selectedIndex].text : "";
	}

	function setSelectValue(select, value) {
		if (!select || !select.options) {
			return;
		}
		for (var index = 0; index < select.options.length; index += 1) {
			if (select.options[index].value === value) {
				select.selectedIndex = index;
				return;
			}
		}
	}

	function currentPurchaseType() {
		var header = headerNode();
		return attr(header, "PurchaseType") || selectedValue("cmbPurType");
	}

	function table() {
		return document.getElementById("tblItemDet");
	}

	function clearTable() {
		var tbl = table();
		if (!tbl) {
			return;
		}
		while (tbl.rows.length > 2) {
			tbl.deleteRow(2);
		}
	}

	function appendCell(row, text, className, align) {
		var cell = row.insertCell();
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		cell.innerHTML = text == null ? "" : String(text);
		return cell;
	}

	function makeInput(name, value, size, className, readOnly) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.value = value == null ? "" : String(value);
		if (size) {
			input.size = size;
		}
		input.className = className || "Formelem";
		input.style.textAlign = "right";
		input.readOnly = !!readOnly;
		return input;
	}

	function makeHidden(name, value) {
		var input = document.createElement("input");
		input.type = "hidden";
		input.name = name;
		input.value = value == null ? "" : String(value);
		return input;
	}

	function itemKey(item) {
		return attr(item, "ClassificationCode") + "Z" + attr(item, "ItemCode") + "Z" + attr(item, "EntryNo");
	}

	function itemDeleteKey(item) {
		return "A" + attr(item, "ClassificationCode") + "A" + attr(item, "ItemCode") + "A" + attr(item, "EntryNo");
	}

	function shouldShowOrgInvoiceColumn() {
		return upper(selectedValue("hFlag")) === "TRUE";
	}

	function renderItemRow(item, serial, singlePurchaseType) {
		var row = table().insertRow(table().rows.length);
		var key = itemKey(item);
		var deleteKey = itemDeleteKey(item);
		var classCode = attr(item, "ClassificationCode");
		var itemCode = attr(item, "ItemCode");
		var entryNo = attr(item, "EntryNo");
		var qty = attr(item, "Qty");
		var rate = attr(item, "Rate");
		var disPer = attr(item, "DisPer");
		var disAmount = attr(item, "DisAmount");
		var nettBasic = attr(item, "NettBasic");
		var basic = numberValue(qty) * numberValue(rate);
		var deleteCheckbox = document.createElement("input");
		var vatCheckbox = document.createElement("input");
		var qtyInput;
		var rateInput;
		var disInput;

		appendCell(row, serial, "ExcelSerial", "center")
			.appendChild(makeHidden("txtSer" + deleteKey, serial));

		deleteCheckbox.type = "checkbox";
		deleteCheckbox.name = "chkDelete" + deleteKey;
		deleteCheckbox.className = "Formelem";
		appendCell(row, "", "ExcelDisplayCell", "center").appendChild(deleteCheckbox);

		appendCell(row, attr(item, "ItmDescription"), "ExcelDisplayCell", "left");
		appendCell(row, attr(item, "UomDesc") || attr(item, "Uom"), "ExcelDisplayCell", "center");

		qtyInput = makeInput("txtInvQty" + key, qty, 10, "Formelem", false);
		qtyInput.onblur = function () { updateAmount(singlePurchaseType ? "SP" : "MP", "Q", itemCode, classCode, entryNo, qtyInput); };
		appendCell(row, "", "ExcelFieldCell", "right").appendChild(qtyInput);

		rateInput = makeInput("txtInvRate" + key, formatNumber(rate, 5), 11, "Formelem", false);
		rateInput.onblur = function () { updateAmount(singlePurchaseType ? "SP" : "MP", "R", itemCode, classCode, entryNo, rateInput); };
		var rateCell = appendCell(row, "", "ExcelFieldCell", "right");
		rateCell.appendChild(rateInput);
		var rateLink = document.createElement("a");
		rateLink.href = "#";
		rateLink.onclick = function () {
			window.SetRateUOM(classCode, itemCode, entryNo, attr(item, "Uom"), singlePurchaseType ? "SP" : "MP", "Y");
			return false;
		};
		rateLink.innerHTML = '&nbsp;<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" width="11" height="11" alt="Select Rate UOM">';
		rateCell.appendChild(rateLink);
		rateCell.appendChild(makeHidden("hRateUOM" + key, attr(item, "RateUOM")));
		rateCell.appendChild(makeHidden("hRatePerQtyUoM" + key, attr(item, "RatePerQtyUoM") || rate));

		appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtValue" + key, formatNumber(basic, 2), 11, "FormelemRead", true));

		disInput = makeInput("txtDisPer" + key, formatNumber(disPer, 2), 6, "Formelem", false);
		disInput.onblur = function () { updateAmount(singlePurchaseType ? "SP" : "MP", "D", itemCode, classCode, entryNo, disInput); };
		appendCell(row, "", "ExcelFieldCell", "right").appendChild(disInput);

		appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtDisAmount" + key, formatNumber(disAmount, 2), 11, "FormelemRead", true));
		appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtNetValue" + key, formatNumber(nettBasic, 2), 11, "FormelemRead", true));

		if (shouldShowOrgInvoiceColumn()) {
			appendCell(row, attr(item, "SuppInvNo") || "", "ExcelDisplayCell", "left");
		}

		appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtTotalItemRate" + (singlePurchaseType ? "0" : attr(item, "PurchaseType")) + "Z" + entryNo, formatNumber(attr(item, "ItemValue") || nettBasic, 2), 11, "FormelemRead", true));

		vatCheckbox.type = "checkbox";
		vatCheckbox.name = "chkVat" + key;
		vatCheckbox.className = "Formelem";
		vatCheckbox.checked = upper(attr(item, "VAT")) === "Y";
		vatCheckbox.onclick = function () {
			item.setAttribute("VAT", vatCheckbox.checked ? "Y" : "N");
		};
		appendCell(row, "", "ExcelDisplayCell", "center").appendChild(vatCheckbox);
	}

	function renderSubtotalRow(prefix, totals) {
		var row = table().insertRow(table().rows.length);
		var labelCell = appendCell(row, "<b>Total</b>", "ExcelSerial", "right");
		labelCell.colSpan = shouldShowOrgInvoiceColumn() ? 6 : 5;
		appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtBasicValue" + prefix, formatNumber(totals.basic, 2), 11, "FormelemRead", true));
		appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtDisValue" + prefix, formatNumber(totals.discount, 2), 11, "FormelemRead", true));
		appendCell(row, "", "ExcelFieldCell", "right");
		appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtAmount" + prefix, formatNumber(totals.nett, 2), 11, "FormelemRead", true));
	}

	function renderTaxRows(taxRoot, prefix) {
		if (!taxRoot) {
			return;
		}
		childElements(taxRoot).forEach(function (tax) {
			var catCode = attr(tax, "CatCode");
			var taxCode = attr(tax, "TaxCode");
			var taxMode = attr(tax, "TaxMode");
			var row;
			var labelCell;
			var perInput;
			if (catCode === "0" && taxCode === "0") {
				return;
			}
			row = table().insertRow(table().rows.length);
			labelCell = appendCell(row, tax.textContent || "", "ExcelDisplayCell", "right");
			labelCell.colSpan = shouldShowOrgInvoiceColumn() ? 7 : 6;

			perInput = makeInput("txtTaxPer" + catCode + taxCode, taxMode === "F" ? "" : attr(tax, "TaxValue"), 6, taxMode === "F" ? "FormelemRead" : "Formelem", taxMode === "F");
			perInput.onchange = window.CalcTax;
			appendCell(row, "", "ExcelFieldCell", "right").appendChild(perInput);
			appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtTaxValue" + catCode + "Z" + taxCode, formatNumber(attr(tax, "TaxAmount"), 2), 11, "FormelemRead", true));
			appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtSubTaxValue" + catCode + taxCode, formatNumber(attr(taxRoot, "SubTotal"), 2), 11, "FormelemRead", true));
		});
		if (prefix) {
			var totalRow = table().insertRow(table().rows.length);
			var totalCell = appendCell(totalRow, "<b>Tax Total</b>", "ExcelSerial", "right");
			totalCell.colSpan = shouldShowOrgInvoiceColumn() ? 8 : 7;
			appendCell(totalRow, "", "ExcelFieldCell", "right").appendChild(makeInput("txtTot" + prefix, formatNumber(attr(taxRoot, "TotalTax"), 2), 11, "FormelemRead", true));
			appendCell(totalRow, "", "ExcelFieldCell", "right").appendChild(makeInput("txtSubTot" + prefix, formatNumber(attr(taxRoot, "SubTotal"), 2), 11, "FormelemRead", true));
		}
	}

	function ensureInvoiceTotalRow() {
		if (field("txtInvValue")) {
			return;
		}
		var row = table().insertRow(table().rows.length);
		var labelCell = appendCell(row, "<b>Invoice Value</b>", "ExcelSerial", "right");
		labelCell.colSpan = shouldShowOrgInvoiceColumn() ? 8 : 7;
		appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtRoundOff", "0.00", 11, "FormelemRead", true));
		appendCell(row, "", "ExcelFieldCell", "right").appendChild(makeInput("txtInvValue", "0.00", 11, "FormelemRead", true));
	}

	function itemTotals(nodes) {
		var totals = { basic: 0, discount: 0, nett: 0 };
		for (var index = 0; index < nodes.length; index += 1) {
			var item = nodes.item(index);
			var qty = numberValue(attr(item, "Qty"));
			var rate = numberValue(attr(item, "Rate"));
			var discount = numberValue(attr(item, "DisAmount"));
			var basic = qty * rate;
			var nett = numberValue(attr(item, "NettBasic")) || (basic - discount);
			totals.basic += basic;
			totals.discount += discount;
			totals.nett += nett;
		}
		return totals;
	}

	function renderSinglePurchaseType() {
		var nodes = itemNodes();
		var totals = itemTotals(nodes);
		var index;
		clearTable();
		for (index = 0; index < nodes.length; index += 1) {
			renderItemRow(nodes.item(index), index + 1, true);
		}
		renderSubtotalRow("", totals);
		renderTaxRows(firstTaxRoot(), "");
		ensureInvoiceTotalRow();
		window.popTax();
		setDeleteEnabled(nodes.length > 0);
	}

	function renderItemwisePurchaseType() {
		var taxes = taxRoots();
		var serial = 1;
		clearTable();
		for (var taxIndex = 0; taxIndex < taxes.length; taxIndex += 1) {
			var taxRoot = taxes.item(taxIndex);
			var purType = attr(taxRoot, "PurchaseType");
			var nodes = itemNodes(purType);
			var heading = table().insertRow(table().rows.length);
			var headingCell = appendCell(heading, "<b>Purchase Type: " + purType + "</b>", "ExcelHeaderCell", "left");
			headingCell.colSpan = shouldShowOrgInvoiceColumn() ? 13 : 12;
			for (var itemIndex = 0; itemIndex < nodes.length; itemIndex += 1) {
				renderItemRow(nodes.item(itemIndex), serial, false);
				serial += 1;
			}
			renderSubtotalRow(purType, itemTotals(nodes));
			renderTaxRows(taxRoot, purType);
		}
		ensureInvoiceTotalRow();
		window.popTax1();
		setDeleteEnabled(itemNodes().length > 0);
	}

	function setDeleteEnabled(enabled) {
		var img = document.images.ImgDeleteIcon || document.getElementsByName("ImgDeleteIcon")[0];
		if (img) {
			img.disabled = !enabled;
		}
	}

	function updateAmount(kind, parameter, itemCode, classCode, entryNo, input) {
		var node = firstNode(selectNodes(invoiceRoot(), "//ItemDetails/Item[@ItemCode=" + xpathLiteral(itemCode) + " and @ClassificationCode=" + xpathLiteral(classCode) + " and @EntryNo=" + xpathLiteral(entryNo) + "]"));
		if (!node || !trim(input.value)) {
			return;
		}
		if (isNaN(Number(input.value))) {
			alert("Enter Numeric Value");
			if (input.select) {
				input.select();
			}
			return;
		}
		if (parameter === "Q") {
			node.setAttribute("Qty", input.value);
		}
		if (parameter === "R") {
			node.setAttribute("Rate", formatNumber(input.value, 5));
			var rateUom = selectedValue("hRateUOM" + classCode + "Z" + itemCode + "Z" + entryNo);
			node.setAttribute("RateUOM", rateUom ? rateUom.split(":")[0] : attr(node, "Uom"));
			node.setAttribute("RatePerQtyUoM", selectedValue("hRatePerQtyUoM" + classCode + "Z" + itemCode + "Z" + entryNo) || formatNumber(input.value, 5));
		}
		if (parameter === "D") {
			node.setAttribute("DisPer", input.value);
		}
		recalculateItem(node);
		if (kind === "SP") {
			renderSinglePurchaseType();
		} else {
			renderItemwisePurchaseType();
		}
	}

	function recalculateItem(node) {
		var qty = numberValue(attr(node, "Qty"));
		var rate = numberValue(attr(node, "Rate"));
		var disPer = numberValue(attr(node, "DisPer"));
		var basic = qty * rate;
		var discount = round((disPer / 100) * basic, 2);
		var nett = round(basic - discount, 2);
		node.setAttribute("Amount", formatNumber(basic, 2));
		node.setAttribute("DisAmount", formatNumber(discount, 2));
		node.setAttribute("NettBasic", formatNumber(nett, 2));
		node.setAttribute("ItemValue", formatNumber(nett, 2));
		node.setAttribute("ItemRate", formatNumber(rate, 5));
	}

	function recalculateAllItems() {
		var nodes = itemNodes();
		for (var index = 0; index < nodes.length; index += 1) {
			recalculateItem(nodes.item(index));
		}
	}

	function removeTaxDetails(root) {
		var nodes = selectNodes(root, "//TaxDetails");
		for (var index = nodes.length - 1; index >= 0; index -= 1) {
			if (nodes.item(index).parentNode) {
				nodes.item(index).parentNode.removeChild(nodes.item(index));
			}
		}
	}

	function appendTaxDetailsFromResponse(responseXML) {
		var responseRoot = xmlRoot(responseXML);
		var root = invoiceRoot();
		childElements(responseRoot).forEach(function (node) {
			root.appendChild(importForInvoice(node));
		});
	}

	function loadTaxDetailsForPurchaseType(purchaseType) {
		var xhr = new XMLHttpRequest();
		var url = "XMLGetTaxDetails.asp?PurType=" + encodeURIComponent(purchaseType) +
			"&ForUnit=" + encodeURIComponent(selectedValue("hOrgID")) +
			"&InvNo=" + encodeURIComponent(selectedValue("hInvNo")) +
			"&RcptNum=" + encodeURIComponent(selectedValue("hRcptNum"));
		xhr.open("GET", url, false);
		xhr.send(null);
		if (xhr.responseXML && xmlRoot(xhr.responseXML)) {
			appendTaxDetailsFromResponse(xhr.responseXML);
		} else if (trim(xhr.responseText)) {
			if (window.TaxFormData && typeof window.TaxFormData.loadXML === "function") {
				window.TaxFormData.loadXML(xhr.responseText);
				appendTaxDetailsFromResponse(window.TaxFormData);
			}
		}
	}

	function updateHeaderPurchaseType(value) {
		var header = headerNode();
		if (header) {
			header.setAttribute("PurchaseType", value);
		}
	}

	function updateItemsPurchaseType(value) {
		var nodes = itemNodes();
		for (var index = 0; index < nodes.length; index += 1) {
			nodes.item(index).setAttribute("PurchaseType", value);
		}
	}

	function calculateTaxAmount(taxRoot, formula, basicValue, displayValue, percentage) {
		var parts = trim(formula).split(",");
		var taxAmount = 0;
		var startIndex = 0;
		if (trim(parts[0]) === "BV") {
			taxAmount = numberValue(basicValue);
			startIndex = 1;
		} else if (trim(parts[0]) === "BD") {
			taxAmount = numberValue(displayValue);
			startIndex = 1;
		}
		for (var index = startIndex; index < parts.length; index += 1) {
			var token = trim(parts[index]).split("#");
			var node = taxNodeByCodes(taxRoot, token[0], token[1]);
			if (node) {
				taxAmount += numberValue(attr(node, "TaxAmount"));
			}
		}
		return trim(percentage) ? taxAmount * (numberValue(percentage) / 100) : taxAmount;
	}

	function recalculateTaxRoot(taxRoot) {
		var purType = attr(taxRoot, "PurchaseType");
		var totals = itemTotals(purType === "0" || purType === "" ? itemNodes() : itemNodes(purType));
		var displayTotal = totals.nett;
		var totalTax = 0;
		childElements(taxRoot).forEach(function (taxNode) {
			var catCode = attr(taxNode, "CatCode");
			var taxCode = attr(taxNode, "TaxCode");
			var mode = attr(taxNode, "TaxMode");
			var tax = 0;
			if (catCode === "0" && taxCode === "0") {
				return;
			}
			if (mode === "P") {
				tax = calculateTaxAmount(taxRoot, attr(taxNode, "TaxFormula"), totals.basic, totals.nett, attr(taxNode, "TaxValue"));
			} else if (mode === "K") {
				var packs = selectNodes(taxRoot, ".//Tax[@CatCode=" + xpathLiteral(catCode) + " and @TaxCode=" + xpathLiteral(taxCode) + "]/Taxpack");
				for (var packIndex = 0; packIndex < packs.length; packIndex += 1) {
					tax += numberValue(packs.item(packIndex).attributes.item(3) && packs.item(packIndex).attributes.item(3).nodeValue);
				}
			} else {
				tax = numberValue(attr(taxNode, "TaxValue"));
			}
			if (attr(taxNode, "Rndoff") === "1") {
				tax = typeof window.RndOff === "function" ? window.RndOff(tax) : Math.round(tax);
			}
			taxNode.setAttribute("TaxAmount", formatNumber(tax, 2));
			totalTax += numberValue(tax);
		});
		taxRoot.setAttribute("Basicvalue", formatNumber(totals.basic, 2));
		taxRoot.setAttribute("NettValue", formatNumber(totals.nett, 2));
		taxRoot.setAttribute("TotalTax", formatNumber(totalTax, 2));
		taxRoot.setAttribute("SubTotal", formatNumber(displayTotal + totalTax, 2));
		return displayTotal + totalTax;
	}

	function recalculateInvoiceTotal() {
		var taxes = taxRoots();
		var invoiceValue = 0;
		var rounded;
		var roundOff;
		for (var index = 0; index < taxes.length; index += 1) {
			invoiceValue += recalculateTaxRoot(taxes.item(index));
		}
		if (!taxes.length) {
			invoiceValue = itemTotals(itemNodes()).nett;
		}
		rounded = typeof window.RndOff === "function" ? window.RndOff(invoiceValue) : Math.round(invoiceValue);
		roundOff = round(rounded - invoiceValue, 2);
		setField("txtRoundOff", formatNumber(roundOff, 2));
		setField("txtInvValue", formatNumber(rounded, 2));
		var header = headerNode();
		if (header) {
			header.setAttribute("InvValue", formatNumber(rounded, 2));
			header.setAttribute("RoundOff", formatNumber(roundOff, 2));
		}
		return rounded;
	}

	function buildBlankInvoiceFromOutData() {
		var root = invoiceRoot();
		var outRoot = xmlRoot(window.OutData);
		var invoiceHeader;
		var header;
		var itemDetails;
		var entryNo = 0;
		removeChildren(root);

		invoiceHeader = createElement("InvoiceHeader");
		header = createElement("Header");
		[
			"OrgID", "Party", "PurchaseType", "Currency", "InvAgainst", "RefNum", "PartyCode", "PartyType",
			"PartySubType", "CurrencyNo", "DespatchMode", "PaymentMode", "PayTerms", "IssueBank",
			"BenificiaryBank", "PricingBasis", "Transporter", "LoadingPort", "DestPort", "Remarks",
			"SuppInvNo", "SuppInvDt", "TransporterFlag", "PoNo", "ConfNum", "InvoiceFlag", "SuppCode",
			"ItemType"
		].forEach(function (name) {
			header.setAttribute(name, "");
		});
		header.setAttribute("OrgID", selectedValue("hOrgID"));
		header.setAttribute("Party", textValue("spnSuppName") || textValue("Span2"));
		header.setAttribute("InvValue", "0");
		header.setAttribute("RoundOff", "0");
		header.setAttribute("InvoiceNumber", selectedValue("hInvNo"));
		invoiceHeader.appendChild(header);
		root.appendChild(invoiceHeader);

		itemDetails = createElement("ItemDetails");
		root.appendChild(itemDetails);

		childElements(outRoot).forEach(function (node) {
			if (node.nodeName === "Item") {
				entryNo += 1;
				appendBlankItem(itemDetails, {
					itemCode: attr(node, "ItemCode"),
					classCode: attr(node, "ClassCode"),
					name: attr(node, "ItemName"),
					uom: attr(node, "StoresUoM"),
					attributeList: attr(node, "AttributeList")
				}, entryNo);
			}
			if (node.nodeName === "Materials") {
				childElements(node).forEach(function (entry) {
					if (entry.nodeName === "Entry") {
						entryNo += 1;
						appendBlankItem(itemDetails, {
							itemCode: "0",
							classCode: "0",
							name: attr(entry, "ItemName"),
							uom: "NOS",
							attributeList: ""
						}, entryNo);
					}
				});
			}
		});
	}

	function appendBlankItem(itemDetails, data, entryNo) {
		var item = createElement("Item");
		item.setAttribute("ItemCode", data.itemCode);
		item.setAttribute("ClassificationCode", data.classCode);
		item.setAttribute("ItmDescription", data.name);
		item.setAttribute("Uom", data.uom);
		item.setAttribute("Qty", "0");
		item.setAttribute("Rate", "0");
		item.setAttribute("DisPer", "0");
		item.setAttribute("DisAmount", "0");
		item.setAttribute("NettBasic", "0");
		item.setAttribute("UomDesc", data.uom);
		item.setAttribute("EntryNo", entryNo);
		item.setAttribute("RatePerQtyUoM", "0");
		item.setAttribute("SourceEntryNo", "");
		item.setAttribute("PurchaseType", "");
		item.setAttribute("Amount", "");
		item.setAttribute("ItemValue", "0");
		item.setAttribute("ItemRate", "0");
		item.setAttribute("RateUOM", "");
		item.setAttribute("StockType", "S");
		item.setAttribute("VAT", "");
		item.setAttribute("AttributeList", data.attributeList || "");
		itemDetails.appendChild(item);
	}

	window.SetDate = function () {
		var date = field("ctlDate");
		if (!date) {
			return;
		}
		if (date.SetMinDate) {
			date.SetMinDate(selectedValue("hMinDate"));
		}
		if (date.SetMaxDate) {
			date.SetMaxDate(selectedValue("hMaxDate"));
		}
	};

	window.MinDate = function () {
		var date = field("ctlDate");
		var current = date && date.GetDate ? date.GetDate() : date && date.value;
		var iso = window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate;
		var valueDate = new Date(iso ? iso(current) : current);
		var minDate = new Date(iso ? iso(selectedValue("hMinDate")) : selectedValue("hMinDate"));
		var maxDate = new Date(iso ? iso(selectedValue("hMaxDate")) : selectedValue("hMaxDate"));
		if (valueDate < minDate || valueDate > maxDate) {
			alert("Date Should be within the Financial Year  " + selectedValue("hMinDate") + " to " + selectedValue("hMaxDate"));
			if (date && date.SetDate) {
				date.SetDate(selectedValue("hMaxDate"));
			}
		}
	};

	window.PopulatePartyType = function () {
		if (!selectedValue("hPartyCode")) {
			return;
		}
		var xhr = new XMLHttpRequest();
		xhr.open("GET", "../../Common/PartySubType.asp?ParCode=" + encodeURIComponent(selectedValue("hPartyCode")) + "&OrgCode=" + encodeURIComponent(selectedValue("hOrgID")), false);
		xhr.send(null);
		if (!trim(xhr.responseText)) {
			return;
		}
		if (window.PartySubTypeData && window.PartySubTypeData.loadXML) {
			window.PartySubTypeData.loadXML(xhr.responseText);
		}
		var root = xmlRoot(window.PartySubTypeData || xhr.responseXML);
		var combo = field("cmbPartyType");
		if (!root || !combo) {
			return;
		}
		combo.length = childElements(root).length === 1 ? 0 : 1;
		childElements(root).forEach(function (node) {
			var parts = attr(node, "SubType").split("|");
			combo.options[combo.length] = new Option(node.textContent || "", (parts[1] || "") + "|" + (parts[0] || ""));
		});
	};

	window.init = function () {
		var date = field("ctlDate");
		if (date && date.SetDate) {
			date.SetDate(selectedValue("hSuppInvDate") || new Date());
		}
		var header = headerNode();
		if (header) {
			setField("txtSuppInvNo", attr(header, "SuppInvNo"));
			setField("txtPartyName", textValue("spnSuppName") || textValue("Span2"));
			if (selectedValue("hPSubType") || selectedValue("hPartyType")) {
				setSelectValue(field("cmbPartyType"), selectedValue("hPSubType") + "|" + selectedValue("hPartyType"));
			}
			if (attr(header, "PurchaseType")) {
				setSelectValue(field("cmbPurType"), attr(header, "PurchaseType"));
			}
		}
		setDeleteEnabled(false);
		if (trim(currentPurchaseType()) !== "0") {
			renderSinglePurchaseType();
		} else {
			renderItemwisePurchaseType();
		}
	};

	window.RefType_Click = function () {
		var refTypeSelect = field("selRefName");
		var refType = refTypeSelect && refTypeSelect.options ? refTypeSelect.options[refTypeSelect.selectedIndex].value : "";
		var orgId = selectedValue("hOrgID");
		var partyCode = selectedValue("hPartyCode");

		function finishSelection() {
			var outRoot = xmlRoot(window.OutData);
			if (trim(refType) !== "N" && outRoot && outRoot.hasChildNodes()) {
				childElements(outRoot).forEach(function (refNode) {
					var remarks = attr(refNode, "Remarks");
					if (remarks) {
						var parts = remarks.split("-");
						partyCode = parts[0] || partyCode;
						setText("spnSuppName", parts[1] || "");
					}
					setText("SpnRcptCode", attr(refNode, "ReferenceCode") + " - " + attr(refNode, "ReferenceDate"));
					setField("hOrgID", orgId);
					setField("hPartyCode", partyCode);
					setField("hRecNo", attr(refNode, "ReferenceNo"));
					setField("hRefType", refType);
					setField("hRefDate", attr(refNode, "ReferenceDate"));
				});
				window.PopulatePartyType();
				var xhr = new XMLHttpRequest();
				xhr.open("GET", "InvItemPopulate.asp?RefType=" + encodeURIComponent(refType) + "&hRecNo=" + encodeURIComponent(selectedValue("hRecNo")) + "&OrgID=" + encodeURIComponent(orgId) + "&InvNo=" + encodeURIComponent(selectedValue("hInvNo")), false);
				xhr.send(null);
				if (trim(xhr.responseText)) {
					loadInvoiceXml(xhr.responseText);
				}
			} else {
				if (!textValue("spnSuppName") && !textValue("Span2")) {
					alert("Select the Party");
					return;
				}
				buildBlankInvoiceFromOutData();
			}
			window.getTaxDet();
		}

		if (typeof window.RefTypeSelection === "function") {
			window.RefTypeSelection(refType, orgId, partyCode, "N", 1, "N", 0, "", finishSelection);
		} else {
			finishSelection();
		}
	};

	window.AddItem = renderSinglePurchaseType;
	window.additem = renderSinglePurchaseType;
	window.PurchaseTypeWise = renderItemwisePurchaseType;

	window.SetRateUOM = function (classCode, itemCode, entryNo, uom, purchaseTypeNature, showWindow) {
		var key = classCode + "Z" + itemCode + "Z" + entryNo;
		var rateUomField = field("hRateUOM" + key);
		var ratePerQtyField = field("hRatePerQtyUoM" + key);
		var rateField = field("txtInvRate" + key);
		var rateUom = rateUomField && rateUomField.value || uom;
		var ratePerQty = ratePerQtyField && ratePerQtyField.value || rateField && rateField.value || "0";
		function applyResult(root) {
			if (root && attr(root, "RateUOM")) {
				rateUom = attr(root, "RateUOM");
				ratePerQty = attr(root, "RatePerQtyUoM");
			}
			if (rateUomField) {
				rateUomField.value = rateUom;
			}
			if (ratePerQtyField) {
				ratePerQtyField.value = ratePerQty;
			}
			if (rateField) {
				rateField.value = formatNumber(ratePerQty, 5);
				updateAmount(upper(purchaseTypeNature) === "SP" ? "SP" : "MP", "R", itemCode, classCode, entryNo, rateField);
			}
		}
		if (trim(showWindow) === "Y") {
			openDialog(
				"invPurInvEntry_RateUomPop.asp?hOrgID=" + encodeURIComponent(selectedValue("hOrgID")) + "&hClassCode=" + encodeURIComponent(classCode) + "&hItemCode=" + encodeURIComponent(itemCode) + "&hEntNo=" + encodeURIComponent(entryNo) + "&hUOM=" + encodeURIComponent(uom) + "&hRateUOM=" + encodeURIComponent(rateUom) + "&hRatePerQtyUOM=" + encodeURIComponent(ratePerQty),
				"",
				"dialogHeight:250px;dialogWidth:300px;status:no",
				applyResult
			);
		} else {
			applyResult(null);
		}
	};

	window.getQtyUoMRate = function () {};
	window.DisplayAmount = function (parameter, itemCode, classCode, entryNo, input) {
		updateAmount("SP", parameter, itemCode, classCode, entryNo, input);
	};
	window.DisplayPurAmount = function (parameter, itemCode, classCode, entryNo, input) {
		updateAmount("MP", parameter, itemCode, classCode, entryNo, input);
	};

	window.setTaxPercentage = function (catCode, taxCode, input) {
		var tax = taxNodeByCodes(firstTaxRoot(), catCode, taxCode);
		if (tax && !isNaN(Number(input.value))) {
			tax.setAttribute("TaxValue", input.value);
			renderSinglePurchaseType();
		}
	};
	window.setTaxAmount = function (catCode, taxCode, input) {
		var tax = taxNodeByCodes(firstTaxRoot(), catCode, taxCode);
		if (tax && !isNaN(Number(input.value))) {
			tax.setAttribute("TaxValue", input.value);
			tax.setAttribute("TaxAmount", input.value);
			renderSinglePurchaseType();
		}
	};
	window.popTax = function () {
		var tax = firstTaxRoot();
		if (tax) {
			recalculateTaxRoot(tax);
		}
		recalculateInvoiceTotal();
	};
	window.popTax1 = function () {
		var taxes = taxRoots();
		for (var index = 0; index < taxes.length; index += 1) {
			recalculateTaxRoot(taxes.item(index));
		}
		recalculateInvoiceTotal();
	};
	window.CalculateTax = function (formula, basicValue, displayValue, percentage, purchaseType) {
		return calculateTaxAmount(firstTaxRoot(purchaseType), formula, basicValue, displayValue, percentage);
	};
	window.ClearTable = clearTable;

	window.getTaxDet = function () {
		var root = invoiceRoot();
		var purType = selectedValue("cmbPurType");
		removeTaxDetails(root);
		updateHeaderPurchaseType(purType);
		updateItemsPurchaseType(purType);
		loadTaxDetailsForPurchaseType(purType);
		childElements(root).forEach(function (node) {
			if (node.nodeName === "TaxDetails") {
				childElements(node).forEach(function (tax) {
					tax.setAttribute("ItemValue", "0");
				});
			}
		});
		saveInvoiceXml();
		if (purType === "0") {
			renderItemwisePurchaseType();
		} else {
			renderSinglePurchaseType();
		}
	};

	window.showTaxFormPopUp = function () {
		alert("The tax form popup file invPurTaxFormPop.asp is not present in the uploaded Purchase folder.");
	};

	window.ShowPurchaseDet = function () {
		saveInvoiceXml();
		openDialog("popItemPurType.asp", window.InvoiceDet, "dialogHeight:400px;dialogWidth:450px;status:no;help:no", function (root) {
			if (root) {
				window.InvoiceDet.XMLDocument.replaceChild(importForInvoice(root), invoiceRoot());
			}
			renderItemwisePurchaseType();
		});
	};

	window.ShowTaxDet = function (purType) {
		openDialog("popTaxDetails.asp?PurType=" + encodeURIComponent(purType), window.InvoiceDet, "dialogLeft:0px;dialogTop:0Px;dialogHeight:400px;dialogWidth:500px;status:no", function (root) {
			if (root) {
				window.InvoiceDet.XMLDocument.replaceChild(importForInvoice(root), invoiceRoot());
			}
			renderItemwisePurchaseType();
		});
	};

	window.getReferenceDetail = function () {};
	window.displayItemDetail = function () {};

	window.Next_Click = function (mode) {
		var header = headerNode();
		var nodes = itemNodes();
		if (!trim(selectedValue("txtSuppInvNo"))) {
			alert("Enter Supplier Invoice No.");
			if (field("txtSuppInvNo")) { field("txtSuppInvNo").focus(); }
			return false;
		}
		if (field("cmbBillType") && field("cmbBillType").selectedIndex === 0) {
			alert("Select Bill Type ");
			field("cmbBillType").focus();
			return false;
		}
		if (field("cmbPurType") && field("cmbPurType").selectedIndex === 0) {
			alert("Select Purchase Type ");
			field("cmbPurType").focus();
			return false;
		}
		if (field("cmbInvCat") && field("cmbInvCat").selectedIndex === 0) {
			alert("Select Category ");
			field("cmbInvCat").focus();
			return false;
		}
		if (field("cmbPartyType") && field("cmbPartyType").selectedIndex === 0) {
			alert("Select Party Type");
			field("cmbPartyType").focus();
			return false;
		}
		for (var index = 0; index < nodes.length; index += 1) {
			var item = nodes.item(index);
			if (numberValue(attr(item, "Qty")) <= 0) {
				alert("Invoice Quantity should be greater than 0");
				return false;
			}
			if (numberValue(attr(item, "Rate")) <= 0) {
				alert("Invoice Rate should be greater than 0");
				return false;
			}
			if (numberValue(attr(item, "DisPer")) < 0) {
				alert("Discount Percentage should not be less than 0");
				return false;
			}
			if (!attr(item, "RateUOM")) {
				item.setAttribute("RateUOM", attr(item, "Uom"));
			}
			item.setAttribute("RatePerQtyUoM", selectedValue("txtInvRate" + itemKey(item)) || attr(item, "RatePerQtyUoM"));
		}
		recalculateAllItems();
		recalculateInvoiceTotal();
		if (header) {
			header.setAttribute("InvValue", selectedValue("txtInvValue"));
			header.setAttribute("RoundOff", selectedValue("txtRoundOff"));
			header.setAttribute("Remarks", selectedValue("mTextAreaRemarks"));
			header.setAttribute("SuppInvNo", selectedValue("txtSuppInvNo"));
			var date = field("ctlDate");
			header.setAttribute("SuppInvDt", date && date.GetDate ? date.GetDate() : selectedValue("ctlDate"));
			header.setAttribute("InvCategory", selectedValue("cmbInvCat"));
			header.setAttribute("BillType", selectedValue("cmbBillType"));
			var partyParts = selectedValue("cmbPartyType").split("|");
			header.setAttribute("PartyType", partyParts[1] || "");
			header.setAttribute("PartySubType", partyParts[0] || "");
		}
		if (mode === "V") {
			return true;
		}
		saveInvoiceXml();
		form().action = "AmdInvPurInvoiceEntryAccDetails.asp";
		form().submit();
		return true;
	};

	window.CalcItemValueForEachItem = function (itemCode, classCode, entryNo) {
		var node = firstNode(selectNodes(invoiceRoot(), "//ItemDetails/Item[@ItemCode=" + xpathLiteral(itemCode) + " and @ClassificationCode=" + xpathLiteral(classCode) + " and @EntryNo=" + xpathLiteral(entryNo) + "]"));
		if (node) {
			recalculateItem(node);
		}
	};
	window.CalculateItemValue = function () {
		recalculateAllItems();
		recalculateInvoiceTotal();
	};
	window.CalcAmountForLastItem = window.CalculateItemValue;
	window.ItemValueCalculateTax = function (taxRoot, formula, basicValue, displayValue, percentage) {
		return calculateTaxAmount(taxRoot, formula, basicValue, displayValue, percentage);
	};
	window.AssignRoundOffValue = recalculateInvoiceTotal;
	window.RoundOffInv = recalculateInvoiceTotal;
	window.AddRoundOffNode = function (roundOffAmount, invoiceValue) {
		var root = invoiceRoot();
		var tax = firstTaxRoot() || createElement("TaxDetails");
		if (!tax.parentNode) {
			root.appendChild(tax);
		}
		if (!taxNodeByCodes(tax, "0", "0")) {
			var node = createElement("Tax");
			node.setAttribute("CatCode", "0");
			node.setAttribute("TaxCode", "0");
			node.setAttribute("TaxValue", roundOffAmount);
			node.setAttribute("TaxAmount", roundOffAmount);
			node.setAttribute("ItemValue", invoiceValue);
			node.textContent = "ROUND OFF";
			tax.appendChild(node);
		}
	};

	window.ShowPrefDet = function () {
		var headerParent = firstNode(selectNodes(invoiceRoot(), "//InvoiceHeader"));
		var header = headerNode();
		if (!headerParent || !header) {
			return;
		}
		openDialog(
			"InvPurInvoiceEntryPref.asp?Mod1=" + encodeURIComponent(attr(header, "DespatchMode")) +
				"&Mop=" + encodeURIComponent(attr(header, "PaymentMode")) +
				"&IssueBank=" + encodeURIComponent(attr(header, "IssueBank")) +
				"&PayTerm=" + encodeURIComponent(attr(header, "PayTerms")) +
				"&Bop=" + encodeURIComponent(attr(header, "PricingBasis")) +
				"&Transporter=" + encodeURIComponent(attr(header, "Transporter")) +
				"&BenefitBank=" + encodeURIComponent(attr(header, "BenificiaryBank")) +
				"&LoadPort=" + encodeURIComponent(attr(header, "LoadingPort")) +
				"&DestPort=" + encodeURIComponent(attr(header, "DestPort")),
			headerParent,
			"dialogLeft:0px;dialogTop:0Px;dialogHeight:250px;dialogWidth:600px;status:no",
			function () {}
		);
	};

	window.DeleteItems = function () {
		var nodes = itemNodes();
		for (var index = nodes.length - 1; index >= 0; index -= 1) {
			var item = nodes.item(index);
			var checkbox = field("chkDelete" + itemDeleteKey(item));
			if (checkbox && checkbox.checked && item.parentNode) {
				item.parentNode.removeChild(item);
			}
		}
		selectedValue("cmbPurType") === "0" ? renderItemwisePurchaseType() : renderSinglePurchaseType();
	};
	window.DeleteItem = window.DeleteItems;
	window.setValues1 = renderItemwisePurchaseType;
	window.setValues = renderSinglePurchaseType;
	window.CalcTax = function () {
		selectedValue("cmbPurType") === "0" ? renderItemwisePurchaseType() : renderSinglePurchaseType();
	};
	window.CalcTotal = recalculateInvoiceTotal;
	window.Packvaluechange = function () {
		alert("The tax package popup file PurInvTaxPackage.asp is not present in the uploaded Purchase folder.");
	};
	window.ViewItemValues = function () {
		alert("The item value popup file InvPurItemValueViewPop.asp is not present in the uploaded Purchase folder.");
	};
	window.SelectAll = function () {
		var checked = !!(field("ChkAll") && field("ChkAll").checked);
		var nodes = itemNodes();
		for (var index = 0; index < nodes.length; index += 1) {
			var checkbox = field("chkVat" + itemKey(nodes.item(index)));
			if (checkbox) {
				checkbox.checked = checked;
				nodes.item(index).setAttribute("VAT", checked ? "Y" : "N");
			}
		}
	};
	window.showInvoiceEntryPop = function () {
		alert("Original supplier invoice selection popup is not available in this uploaded set.");
	};
})(window, document);
