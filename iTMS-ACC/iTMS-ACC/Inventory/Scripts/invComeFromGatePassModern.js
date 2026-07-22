(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms.Formname || document.forms[0];
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function controls(name) {
		var frm = form();
		var direct;
		var elements;
		if (!frm || !frm.elements) {
			return null;
		}
		direct = frm.elements[name];
		if (direct) {
			return direct;
		}
		elements = frm.elements;
		for (var i = 0; i < elements.length; i += 1) {
			if (String(elements[i].name || elements[i].id).toLowerCase() === String(name).toLowerCase()) {
				return elements[i];
			}
		}
		return null;
	}

	function field(name) {
		var found = controls(name);
		if (found && found.length && !found.tagName) {
			return found[0];
		}
		return found || document.getElementById(name) || null;
	}

	function fieldValue(name) {
		var item = field(name);
		return item ? item.value || "" : "";
	}

	function setFieldValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function itemField(prefix, classCode, itemCode, entryNo) {
		return field(prefix + classCode + "Z" + itemCode + "Z" + entryNo) ||
			field(prefix + classCode + "Z" + itemCode);
	}

	function deleteField(classCode, itemCode, entryNo) {
		return field("chkDeleteA" + classCode + "A" + itemCode + "A" + entryNo);
	}

	function serialField(classCode, itemCode, entryNo) {
		return field("txtSerA" + classCode + "A" + itemCode + "A" + entryNo);
	}

	function vatField(classCode, itemCode, entryNo) {
		return field("chkVat" + classCode + "Z" + itemCode + "Z" + entryNo) ||
			field("chkVat" + classCode + "Z" + itemCode);
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function isNumeric(value) {
		return trim(value) !== "" && isFinite(Number(String(value).replace(/,/g, "")));
	}

	function roundTo(value, decimals) {
		var factor = Math.pow(10, decimals || 0);
		return Math.round(toNumber(value) * factor) / factor;
	}

	function formatNumber(value, decimals) {
		var places = decimals == null ? 2 : decimals;
		return roundTo(value, places).toFixed(places);
	}

	function roundInvoice(value) {
		if (typeof window.RndOff === "function") {
			return toNumber(window.RndOff(value));
		}
		return Math.round(toNumber(value));
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

	function dateDiffDays(start, end) {
		var startDate = parseDate(start);
		var endDate = parseDate(end);
		if (!startDate || !endDate) {
			return 0;
		}
		return Math.floor((endDate.getTime() - startDate.getTime()) / 86400000);
	}

	function controlDate(name) {
		var item = field(name);
		if (!item) {
			return "";
		}
		if (typeof item.GetDate === "function") {
			return item.GetDate();
		}
		if (typeof item.getDate === "function") {
			return item.getDate();
		}
		return item.value || "";
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
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement ||
			value._doc && value._doc.documentElement || value;
	}

	function invoiceDoc() {
		return xmlDocument(xmlIsland("InvoiceDet"));
	}

	function invoiceRoot() {
		return xmlRoot(xmlIsland("InvoiceDet"));
	}

	function serializeXml(value) {
		var doc = value && value.nodeType === 9 ? value : xmlDocument(value);
		return doc ? new XMLSerializer().serializeToString(doc) : "";
	}

	function loadXmlIntoIsland(name, text) {
		var island = xmlIsland(name);
		if (island && typeof island.loadXML === "function") {
			return island.loadXML(text);
		}
		return false;
	}

	function importNode(doc, node) {
		if (!doc || !node) {
			return null;
		}
		return doc.importNode ? doc.importNode(node, true) : node.cloneNode(true);
	}

	function replaceInvoiceRoot(value) {
		var doc = invoiceDoc();
		var root = xmlRoot(value);
		if (!doc || !root) {
			return;
		}
		while (doc.firstChild) {
			doc.removeChild(doc.firstChild);
		}
		doc.appendChild(importNode(doc, root));
	}

	function children(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var nodes = node && node.childNodes || [];
		for (var i = 0; i < nodes.length; i += 1) {
			if (nodes[i].nodeType === 1 && (!wanted || String(nodes[i].nodeName).toLowerCase() === wanted)) {
				result.push(nodes[i]);
			}
		}
		return result;
	}

	function attr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function firstChild(root, name) {
		return children(root, name)[0] || null;
	}

	function headerNode() {
		return firstChild(firstChild(invoiceRoot(), "InvoiceHeader"), "Header");
	}

	function itemDetailsNode() {
		var root = invoiceRoot();
		var doc = invoiceDoc();
		var node = firstChild(root, "ItemDetails");
		if (!node && root && doc) {
			node = doc.createElement("ItemDetails");
			root.appendChild(node);
		}
		return node;
	}

	function items() {
		return children(itemDetailsNode(), "Item");
	}

	function taxDetailsNodes() {
		return children(invoiceRoot(), "TaxDetails");
	}

	function taxNodes(taxRoot) {
		return children(taxRoot, "Tax");
	}

	function taxDetailsForPurchaseType(purchaseType) {
		var nodes = taxDetailsNodes();
		for (var i = 0; i < nodes.length; i += 1) {
			if (trim(attr(nodes[i], "PurchaseType")) === trim(purchaseType)) {
				return nodes[i];
			}
		}
		return trim(purchaseType) === "0" || trim(purchaseType) === "" ? nodes[0] || null : null;
	}

	function itemsForPurchaseType(purchaseType) {
		return items().filter(function (item) {
			return trim(attr(item, "PurchaseType")) === trim(purchaseType);
		});
	}

	function matchingItem(itemCode, classCode, entryNo) {
		var list = items();
		for (var i = 0; i < list.length; i += 1) {
			if (trim(attr(list[i], "ItemCode")) === trim(itemCode) &&
					trim(attr(list[i], "ClassificationCode")) === trim(classCode) &&
					(!entryNo || trim(attr(list[i], "SourceEntryNo")) === trim(entryNo))) {
				return list[i];
			}
		}
		return null;
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

	function selectedIndex(name) {
		var control = field(name);
		return control && control.options ? control.selectedIndex : -1;
	}

	function selectedText(name) {
		var control = field(name);
		if (control && control.options && control.selectedIndex >= 0) {
			return control.options[control.selectedIndex].text;
		}
		return "";
	}

	function setDisabled(name, disabled) {
		var item = field(name) || document.getElementsByName(name)[0];
		if (item) {
			item.disabled = !!disabled;
		}
	}

	function openModal(url, args, features, callback) {
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
		xhr.send(body || "");
		return xhr;
	}

	function saveInvoice(name) {
		return syncPost("XMLSavePur.asp?Mod=PUR&Name=" + encodeURIComponent(name), serializeXml(xmlIsland("InvoiceDet")));
	}

	function createInput(type, name, value, className, size, readOnly) {
		var input = document.createElement("input");
		input.type = type || "text";
		input.name = name;
		if (value != null) {
			input.value = value;
		}
		if (className) {
			input.className = className;
		}
		if (size) {
			input.size = size;
		}
		if (readOnly) {
			input.readOnly = true;
		}
		return input;
	}

	function addCell(row, className, align, width, colSpan) {
		var cell = row.insertCell(-1);
		if (className) {
			cell.className = className;
		}
		if (align) {
			cell.align = align;
		}
		if (width) {
			cell.width = width;
		}
		if (colSpan) {
			cell.colSpan = colSpan;
		}
		return cell;
	}

	function appendHidden(container, name, value) {
		var input = createInput("hidden", name, value || "", "", 0, false);
		container.appendChild(input);
		return input;
	}

	function makeIcon(alt, handler) {
		var link = document.createElement("a");
		var img = document.createElement("img");
		img.border = "0";
		img.src = "../../assets/images/iTMS%20Icons/EntryIcon.gif";
		img.alt = alt;
		img.style.cursor = "pointer";
		img.width = 15;
		img.height = 15;
		link.href = "#";
		link.onclick = function () {
			handler();
			return false;
		};
		link.appendChild(img);
		return link;
	}

	function table() {
		return document.getElementById("tblItemDet");
	}

	function currentPurchaseType() {
		return selectedValue("cmbPurType");
	}

	function renderItemRow(item, serial, purchaseNature) {
		var row = table().insertRow(-1);
		var itemCode = attr(item, "ItemCode");
		var classCode = attr(item, "ClassificationCode");
		var entryNo = attr(item, "SourceEntryNo");
		var uom = attr(item, "Uom");
		var desc = attr(item, "ItmDescription");
		var qty = attr(item, "Qty") || "0";
		var rate = formatNumber(attr(item, "Rate"), 5);
		var disPer = attr(item, "DisPer") || "0";
		var disAmount = attr(item, "DisAmount") || "0";
		var netBasic = attr(item, "NettBasic") || "0";
		var cell;
		var input;

		cell = addCell(row, "ExcelDisplayCell", "center");
		cell.appendChild(createInput("text", "txtSerA" + classCode + "A" + itemCode + "A" + entryNo, serial, "FormelemRead", 1, true));

		cell = addCell(row, "ExcelDisplayCell", "", "10");
		cell.appendChild(createInput("checkbox", "chkDeleteA" + classCode + "A" + itemCode + "A" + entryNo, "Y", "FormElem"));

		cell = addCell(row, "ExcelDisplayCell", "left");
		input = document.createElement("a");
		input.className = "ExcelDisplayLink";
		input.href = "#";
		input.textContent = desc;
		input.onclick = function () {
			window.GetAddiDesc(classCode, itemCode, entryNo);
			return false;
		};
		cell.appendChild(input);

		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		cell.appendChild(document.createTextNode(uom));
		appendHidden(cell, "hRateUOM" + classCode + "Z" + itemCode + "Z" + entryNo, attr(item, "RateUOM"));
		appendHidden(cell, "hRatePerQtyUoM" + classCode + "Z" + itemCode + "Z" + entryNo, attr(item, "RatePerQtyUoM"));

		cell = addCell(row, "ExcelInputCell", "center", "10");
		input = createInput("text", "txtInvQty" + classCode + "Z" + itemCode + "Z" + entryNo, qty, "Formelem", 10);
		input.style.textAlign = "right";
		input.onblur = function () {
			if (purchaseNature === "MP") {
				window.DisplayPurAmount("Q", itemCode, classCode, entryNo, this);
			} else {
				window.DisplayAmount("Q", itemCode, classCode, entryNo, this);
			}
		};
		cell.appendChild(input);

		cell = addCell(row, "ExcelInputCell", "center", "100");
		input = createInput("text", "txtInvRate" + classCode + "Z" + itemCode + "Z" + entryNo, rate, "Formelem", 11);
		input.style.textAlign = "right";
		input.onblur = function () {
			if (purchaseNature === "MP") {
				window.DisplayPurAmount("R", itemCode, classCode, entryNo, this);
			} else {
				window.DisplayAmount("R", itemCode, classCode, entryNo, this);
			}
		};
		cell.appendChild(input);
		cell.appendChild(document.createTextNode("  "));
		cell.appendChild(makeIcon("Select Rate UOM", function () {
			window.SetRateUOM(classCode, itemCode, uom, purchaseNature, entryNo);
		}));

		cell = addCell(row, "FormelemRead", "center", "10");
		input = createInput("text", "txtValue" + classCode + "Z" + itemCode + "Z" + entryNo, "", "FormelemRead", 11, true);
		input.style.textAlign = "right";
		cell.appendChild(input);

		cell = addCell(row, "ExcelInputCell", "center", "10");
		input = createInput("text", "txtDisPer" + classCode + "Z" + itemCode + "Z" + entryNo, disPer, "Formelem", 6);
		input.style.textAlign = "right";
		input.onblur = function () {
			if (purchaseNature === "MP") {
				window.DisplayPurAmount("D", itemCode, classCode, entryNo, this);
			} else {
				window.DisplayAmount("D", itemCode, classCode, entryNo, this);
			}
		};
		cell.appendChild(input);

		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		input = createInput("text", "txtDisAmount" + classCode + "Z" + itemCode + "Z" + entryNo, disAmount, "FormelemRead", 11, true);
		input.style.textAlign = "right";
		cell.appendChild(input);

		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		input = createInput("text", "txtNetValue" + classCode + "Z" + itemCode + "Z" + entryNo, netBasic, "FormelemRead", 11, true);
		input.style.textAlign = "right";
		cell.appendChild(input);

		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		input = createInput("text", "txtItemRateDisplay" + classCode + "Z" + itemCode + "Z" + entryNo, "", "FormelemRead", 11, true);
		input.style.textAlign = "right";
		cell.appendChild(input);

		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		input = createInput("checkbox", "chkVat" + classCode + "Z" + itemCode + "Z" + entryNo, "Y", "Formelem");
		input.checked = trim(attr(item, "VAT")).toUpperCase() === "Y";
		cell.appendChild(input);
	}

	function renderRoundOffRows(itemWise) {
		var row;
		var cell;
		var yes;
		var no;
		var input;

		row = table().insertRow(-1);
		cell = addCell(row, "ExcelSerial", "right", "", 9);
		cell.appendChild(document.createElement("b")).appendChild(document.createTextNode("Rounded Off"));
		cell.appendChild(document.createTextNode("("));
		yes = createInput("radio", "rdRndOff", "Y", "Formelem");
		yes.onclick = window.RoundOffInv;
		cell.appendChild(yes);
		cell.appendChild(document.createTextNode("Yes"));
		no = createInput("radio", "rdRndOff", "N", "Formelem");
		no.onclick = window.RoundOffInv;
		cell.appendChild(no);
		cell.appendChild(document.createTextNode("No)"));
		cell = addCell(row, "ExcelDisplayCell", "right", "10");
		input = createInput("text", "txtRoundOff", "", "FormelemRead", 11, true);
		input.style.textAlign = "right";
		input.onblur = window.AssignRoundOffValue;
		cell.appendChild(input);
		addCell(row, "ExcelDisplayCell", "right", "10");

		row = table().insertRow(-1);
		cell = addCell(row, "ExcelSerial", "right", "", 9);
		cell.appendChild(document.createElement("b")).appendChild(document.createTextNode("Invoice Value"));
		cell = addCell(row, "ExcelDisplayCell", "right", "10");
		input = createInput("text", "txtInvValue", "", "FormelemRead", 11, true);
		input.style.textAlign = "right";
		cell.appendChild(input);
		cell = addCell(row, "ExcelDisplayCell", "right", "10");
		if (itemWise) {
			input = createInput("text", "mTxtTotal", "", "FormelemRead", 11, true);
			input.style.textAlign = "right";
			cell.appendChild(input);
		}
	}

	function renderSingleTotals(purchaseType) {
		var row = table().insertRow(-1);
		var cell = addCell(row, "ExcelSerial", "right", "", 6);
		var checkbox = createInput("checkbox", "chkConsolidated", "Y", "Formelem");
		checkbox.onclick = window.chkConsolidated_onClick;
		cell.appendChild(document.createTextNode("[ "));
		cell.appendChild(checkbox);
		cell.appendChild(document.createTextNode("Enter Consolidated Value  ] "));
		cell.appendChild(document.createElement("b")).appendChild(document.createTextNode("Total"));

		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		cell.appendChild(createInput("text", "txtBasicValue", "", "FormelemRead", 11, true));
		addCell(row, "ExcelDisplayCell", "", "10");
		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		cell.appendChild(createInput("text", "txtDisValue", "", "FormelemRead", 11, true));
		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		cell.appendChild(createInput("text", "txtAmount", "", "FormelemRead", 11));
		cell.firstChild.style.textAlign = "right";
		cell.firstChild.onchange = window.Amount_onChange;
		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		cell.appendChild(createInput("text", "txtTotalItemRate" + purchaseType, "", "FormelemRead", 11, true));
		addCell(row, "ExcelDisplayCell", "", "10");

		taxDetailsNodes().forEach(function (taxRoot) {
			taxNodes(taxRoot).forEach(function (tax) {
				var taxMode = attr(tax, "TaxMode");
				var catCode = attr(tax, "CatCode");
				var taxCode = attr(tax, "TaxCode");
				var taxVal = taxMode === "F" ? attr(tax, "TaxValue") : "";
				var taxPer = taxMode === "F" ? "" : attr(tax, "TaxValue");
				var input;
				row = table().insertRow(-1);
				cell = addCell(row, "ExcelSerial", "right", "", 8);
				cell.textContent = tax.textContent || attr(tax, "TaxName") || "";
				cell = addCell(row, "ExcelFieldCell", "center", "10");
				if (taxMode !== "F") {
					input = createInput("text", "txtTaxPer" + catCode + taxCode, taxPer, "Formelem", 6);
					input.style.textAlign = "right";
					input.onblur = function () {
						window.setTaxPercentage(catCode, taxCode, this);
					};
					cell.appendChild(input);
					cell.appendChild(document.createTextNode(" %"));
				}
				cell = addCell(row, "ExcelFieldCell", "center", "10");
				input = createInput("text", "txtTaxValue" + catCode + taxCode, taxVal, taxMode === "F" ? "Formelem" : "FormelemRead", 11, taxMode !== "F");
				input.style.textAlign = "right";
				input.onblur = function () {
					window.setTaxAmount(catCode, taxCode, this);
				};
				cell.appendChild(input);
				cell = addCell(row, "ExcelAverageCell", "center", "10");
				input = createInput("text", "txtSubTaxValue" + catCode + taxCode, "", "FormelemRead", 11, true);
				input.style.textAlign = "right";
				cell.appendChild(input);
				addCell(row, "ExcelAverageCell", "", "10");
			});
		});
	}

	function renderItemWiseTotals(purchaseType) {
		var row = table().insertRow(-1);
		var cell = addCell(row, "ExcelSerial", "right", "", 6);
		var input;
		cell.appendChild(document.createElement("b")).appendChild(document.createTextNode("Total"));
		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		cell.appendChild(createInput("text", "txtBasicValue" + purchaseType, "", "FormelemRead", 11, true));
		addCell(row, "ExcelDisplayCell", "", "10");
		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		cell.appendChild(createInput("text", "txtDisValue" + purchaseType, "", "FormelemRead", 11, true));
		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		cell.appendChild(createInput("text", "txtAmount" + purchaseType, "", "FormelemRead", 11, true));
		cell = addCell(row, "ExcelDisplayCell", "center", "10");
		cell.appendChild(createInput("text", "txtTotalItemRate" + purchaseType, "", "FormelemRead", 11, true));

		row = table().insertRow(-1);
		cell = addCell(row, "ExcelSerial", "right", "", 9);
		cell.appendChild(document.createTextNode("Total Tax Value "));
		cell.appendChild(makeIcon("View Tax Details", function () {
			window.ShowTaxDet(purchaseType);
		}));
		cell = addCell(row, "ExcelDisplayCell", "right", "10");
		input = createInput("text", "txtTot" + purchaseType, "", "FormelemRead", 11, true);
		input.style.textAlign = "right";
		cell.appendChild(input);
		cell = addCell(row, "ExcelDisplayCell", "right", "10");
		input = createInput("text", "txtTaxDisplay" + purchaseType, "", "FormelemRead", 11, true);
		input.style.textAlign = "right";
		cell.appendChild(input);

		row = table().insertRow(-1);
		cell = addCell(row, "ExcelSerial", "right", "", 9);
		cell.textContent = "Sub Total";
		cell = addCell(row, "ExcelDisplayCell", "right", "10");
		input = createInput("text", "txtSubTot" + purchaseType, "", "FormelemRead", 11, true);
		input.style.textAlign = "right";
		cell.appendChild(input);
		addCell(row, "ExcelDisplayCell", "right", "10");
	}

	function renderCurrentMode() {
		if (currentPurchaseType() === "0") {
			window.PurchaseTypeWise();
		} else {
			window.AddItem();
		}
	}

	function updateLineAmount(item, classCode, itemCode, entryNo) {
		var qty = toNumber(attr(item, "Qty"));
		var rate = toNumber(attr(item, "Rate"));
		var disPer = toNumber(attr(item, "DisPer"));
		var basic = qty * rate;
		var discount = roundTo((disPer / 100) * basic, 2);
		var net = roundTo(basic - discount, 2);
		var input;
		setAttr(item, "DisAmount", discount);
		setAttr(item, "NettBasic", net);
		input = itemField("txtValue", classCode, itemCode, entryNo);
		if (input) {
			input.value = formatNumber(basic, 2);
		}
		input = itemField("txtDisAmount", classCode, itemCode, entryNo);
		if (input) {
			input.value = formatNumber(discount, 2);
		}
		input = itemField("txtNetValue", classCode, itemCode, entryNo);
		if (input) {
			input.value = formatNumber(net, 2);
		}
	}

	function updateTotalsForItems(list, suffix) {
		var basic = 0;
		var discount = 0;
		var net = 0;
		list.forEach(function (item) {
			var qty = toNumber(attr(item, "Qty"));
			var rate = toNumber(attr(item, "Rate"));
			var dis = toNumber(attr(item, "DisAmount"));
			basic += qty * rate;
			discount += dis;
			net += toNumber(attr(item, "NettBasic"));
		});
		setFieldValue("txtBasicValue" + (suffix || ""), formatNumber(basic, 2));
		setFieldValue("txtDisValue" + (suffix || ""), formatNumber(discount, 2));
		setFieldValue("txtAmount" + (suffix || ""), formatNumber(net, 2));
		return { basic: basic, discount: discount, net: net };
	}

	function calculateFormulaTax(taxRoot, formula, basicValue, netValue, percentage, useItemValue) {
		var parts = String(formula || "").split(",");
		var base = 0;
		var start = 0;
		if (trim(parts[0]) === "BV") {
			base = toNumber(basicValue);
			start = 1;
		} else if (trim(parts[0]) === "BD") {
			base = toNumber(netValue);
			start = 1;
		}
		for (var i = start; i < parts.length; i += 1) {
			var ref = trim(parts[i]).split("#");
			taxNodes(taxRoot).forEach(function (tax) {
				if (trim(attr(tax, "CatCode")) === trim(ref[0]) && trim(attr(tax, "TaxCode")) === trim(ref[1])) {
					base += toNumber(useItemValue ? attr(tax, "ItemValue") : attr(tax, "TaxAmount"));
				}
			});
		}
		return trim(percentage) ? base * (toNumber(percentage) / 100) : base;
	}

	function taxPackValue(tax) {
		var total = 0;
		children(tax, "Taxpack").forEach(function (pack) {
			total += toNumber(pack.attributes && pack.attributes.item(3) ? pack.attributes.item(3).value : 0);
		});
		return total;
	}

	function setRoundOffControls(invoiceAmount) {
		var roundedInvoice = roundInvoice(invoiceAmount);
		var roundedOff = roundTo(roundedInvoice - toNumber(invoiceAmount), 2);
		var radio = controls("rdRndOff");
		if (radio && radio.length) {
			radio[roundedOff !== 0 ? 0 : 1].checked = true;
		}
		if (radio && radio.length && radio[0].checked) {
			setFieldValue("txtRoundOff", formatNumber(roundedOff, 2));
			setFieldValue("txtInvValue", formatNumber(roundedInvoice, 2));
		} else {
			setFieldValue("txtRoundOff", "0.00");
			setFieldValue("txtInvValue", formatNumber(invoiceAmount, 2));
		}
	}

	function updateHeaderTotals() {
		var header = headerNode();
		if (header) {
			setAttr(header, "InvValue", fieldValue("txtInvValue"));
			setAttr(header, "RoundOff", fieldValue("txtRoundOff"));
		}
	}

	function updateTaxRoundOffNode(taxRoot) {
		taxNodes(taxRoot).forEach(function (tax) {
			if (attr(tax, "CatCode") === "0" && attr(tax, "TaxCode") === "0") {
				setAttr(tax, "TaxValue", fieldValue("txtRoundOff"));
				setAttr(tax, "TaxAmount", fieldValue("txtRoundOff"));
			}
		});
	}

	window.GetAddiDesc = function (classCode, itemCode, entryNo) {
		openModal("invPurInvEntry_ItemAddDesc.asp?ClassCode=" + encodeURIComponent(classCode) + "&ItemCode=" + encodeURIComponent(itemCode), xmlIsland("InvoiceDet"), "dialogHeight:200px;dialogWidth:450px;status:no", function (result) {
			var root = xmlRoot(result);
			var item = matchingItem(itemCode, classCode, entryNo);
			var addDesc = attr(root, "ItemAddiDesc");
			if (item && addDesc) {
				setAttr(item, "ItemAddiDesc", addDesc);
			}
		});
		return false;
	};

	window.ClearTable = function () {
		var tbl = table();
		if (!tbl) {
			return false;
		}
		while (tbl.rows.length > 2) {
			tbl.deleteRow(2);
		}
		return false;
	};

	window.AddItem = function () {
		var purchaseType = currentPurchaseType();
		var serial = 1;
		window.ClearTable();
		items().forEach(function (item) {
			if (purchaseType && purchaseType !== "0") {
				setAttr(item, "PurchaseType", purchaseType);
			}
			renderItemRow(item, serial, "SP");
			serial += 1;
		});
		renderSingleTotals(purchaseType);
		renderRoundOffRows(false);
		setDisabled("ImgDeleteIcon", false);
		items().forEach(function (item) {
			window.DisplayAmount("R", attr(item, "ItemCode"), attr(item, "ClassificationCode"), attr(item, "SourceEntryNo"), itemField("txtInvRate", attr(item, "ClassificationCode"), attr(item, "ItemCode"), attr(item, "SourceEntryNo")));
		});
		return false;
	};

	window.addItem = window.AddItem;

	window.PurchaseTypeWise = function () {
		var serial = 1;
		window.ClearTable();
		taxDetailsNodes().forEach(function (taxRoot) {
			var purchaseType = attr(taxRoot, "PurchaseType");
			var groupItems = itemsForPurchaseType(purchaseType);
			groupItems.forEach(function (item) {
				renderItemRow(item, serial, "MP");
				serial += 1;
			});
			if (groupItems.length > 0) {
				renderItemWiseTotals(purchaseType);
			}
		});
		if (taxDetailsNodes().length === 0) {
			items().forEach(function (item) {
				renderItemRow(item, serial, "MP");
				serial += 1;
			});
		}
		renderRoundOffRows(true);
		setDisabled("ImgDeleteIcon", false);
		items().forEach(function (item) {
			window.DisplayPurAmount("R", attr(item, "ItemCode"), attr(item, "ClassificationCode"), attr(item, "SourceEntryNo"), itemField("txtInvRate", attr(item, "ClassificationCode"), attr(item, "ItemCode"), attr(item, "SourceEntryNo")));
		});
		return false;
	};

	window.Amount_onChange = function () {
		var amount = fieldValue("txtAmount");
		if (field("chkConsolidated") && field("chkConsolidated").checked) {
			items().forEach(function (item) {
				setAttr(item, "Rate", amount);
				setAttr(item, "RatePerQtyUoM", amount);
				setAttr(item, "Amount", amount);
				setAttr(item, "ItemValue", amount);
				setAttr(item, "ItemRate", amount);
				setAttr(item, "NettBasic", amount);
			});
		}
		window.CalculateItemValue(currentPurchaseType());
		window.popTax();
		return false;
	};

	window.chkConsolidated_onClick = function () {
		var checked = field("chkConsolidated") && field("chkConsolidated").checked;
		items().forEach(function (item) {
			var classCode = attr(item, "ClassificationCode");
			var itemCode = attr(item, "ItemCode");
			var entryNo = attr(item, "SourceEntryNo");
			var rate = itemField("txtInvRate", classCode, itemCode, entryNo);
			var discount = itemField("txtDisPer", classCode, itemCode, entryNo);
			if (rate) {
				rate.disabled = checked;
				if (checked) {
					rate.value = "0";
				}
			}
			if (discount) {
				discount.disabled = checked;
				if (checked) {
					discount.value = "0";
				}
			}
			if (checked) {
				["txtValue", "txtNetValue", "txtItemRateDisplay"].forEach(function (prefix) {
					var control = itemField(prefix, classCode, itemCode, entryNo);
					if (control) {
						control.value = "0";
					}
				});
			}
		});
		if (field("txtAmount")) {
			field("txtAmount").className = checked ? "FormElem" : "FormElemRead";
			if (!checked) {
				field("txtAmount").value = "0";
			}
		}
		window.CalculateItemValue(currentPurchaseType());
		return false;
	};

	window.SetRateUOM = function (classCode, itemCode, uom, purchaseNature, entryNo) {
		var orgId = fieldValue("hOrgID");
		var rateUomControl = itemField("hRateUOM", classCode, itemCode, entryNo);
		var ratePerControl = itemField("hRatePerQtyUoM", classCode, itemCode, entryNo);
		var rateUom = trim(rateUomControl && rateUomControl.value) || uom;
		var ratePer = trim(ratePerControl && ratePerControl.value);
		var url = "invPurInvEntry_RateUomPop.asp?hOrgID=" + encodeURIComponent(orgId) +
			"&hClassCode=" + encodeURIComponent(classCode) +
			"&hItemCode=" + encodeURIComponent(itemCode) +
			"&hUOM=" + encodeURIComponent(uom) +
			"&hRateUOM=" + encodeURIComponent(rateUom) +
			"&hRatePerQtyUOM=" + encodeURIComponent(ratePer);
		openModal(url, "", "dialogHeight:250px;dialogWidth:300px;status:no", function (result) {
			var root = xmlRoot(result);
			var selectedRateUom = attr(root, "RateUOM") || uom;
			var selectedRatePer = attr(root, "RatePerQtyUoM") || fieldValue("txtInvRate" + classCode + "Z" + itemCode + "Z" + entryNo);
			if (rateUomControl) {
				rateUomControl.value = selectedRateUom;
			}
			if (ratePerControl) {
				ratePerControl.value = selectedRatePer;
			}
			window.getQtyUoMRate(orgId, classCode, itemCode, rateUomControl, uom, entryNo);
			if (String(purchaseNature).toUpperCase() === "SP") {
				window.DisplayAmount("R", itemCode, classCode, entryNo, itemField("txtInvRate", classCode, itemCode, entryNo));
			} else {
				window.DisplayPurAmount("R", itemCode, classCode, entryNo, itemField("txtInvRate", classCode, itemCode, entryNo));
			}
		});
		return false;
	};

	window.getQtyUoMRate = function (orgId, classCode, itemCode, obj, uom, entryNo) {
		var qtyUom = trim(String(uom || "").split(":")[0]);
		var rateField = itemField("txtInvRate", classCode, itemCode, entryNo);
		var rateUom = trim(obj && obj.value || "");
		var rate = rateField ? toNumber(rateField.value) : 0;
		var converted = rate;
		if (rateUom) {
			rateUom = trim(rateUom.split(":")[0]);
		}
		if (trim(classCode) !== "0" && trim(classCode).toUpperCase() !== "TEMP" && rateUom && rateUom !== qtyUom && typeof window.getRatePerQtyUoM === "function") {
			converted = window.getRatePerQtyUoM(orgId, classCode, itemCode, qtyUom, rateUom, rate);
		}
		if (rateField) {
			rateField.value = formatNumber(converted, 5);
		}
		return converted;
	};

	function displayAmountCommon(kind, itemCode, classCode, entryNo, objText, itemWise) {
		var item = matchingItem(itemCode, classCode, entryNo);
		var purchaseType;
		if (!item || !objText) {
			return false;
		}
		if (!isNumeric(objText.value)) {
			alert("Enter Numeric Value");
			if (objText.select) {
				objText.select();
			}
			return false;
		}
		if (kind === "Q") {
			setAttr(item, "Qty", objText.value);
		} else if (kind === "R") {
			setAttr(item, "Rate", formatNumber(objText.value, 5));
			if (itemField("hRatePerQtyUoM", classCode, itemCode, entryNo) && trim(itemField("hRatePerQtyUoM", classCode, itemCode, entryNo).value)) {
				setAttr(item, "RatePerQtyUoM", itemField("hRatePerQtyUoM", classCode, itemCode, entryNo).value);
				setAttr(item, "RateUOM", trim((itemField("hRateUOM", classCode, itemCode, entryNo).value || "").split(":")[0]));
			} else {
				setAttr(item, "RatePerQtyUoM", formatNumber(objText.value, 5));
				setAttr(item, "RateUOM", attr(item, "Uom"));
			}
		} else if (kind === "D") {
			setAttr(item, "DisPer", objText.value);
		}
		updateLineAmount(item, classCode, itemCode, entryNo);
		purchaseType = itemWise ? attr(item, "PurchaseType") : currentPurchaseType();
		if (itemWise) {
			updateTotalsForItems(itemsForPurchaseType(purchaseType), purchaseType);
			window.popTax1();
		} else {
			if (!purchaseType || purchaseType === "0") {
				updateTotalsForItems(items(), "");
				return false;
			}
			updateTotalsForItems(items(), "");
			window.popTax();
		}
		window.CalcItemValueForEachItem(itemCode, classCode);
		if (kind === "R") {
			objText.value = formatNumber(objText.value, 5);
		} else if (kind === "D") {
			objText.value = formatNumber(objText.value, 2);
		}
		return false;
	}

	window.DisplayAmount = function (kind, itemCode, classCode, entryNo, objText) {
		return displayAmountCommon(kind, itemCode, classCode, entryNo, objText, false);
	};

	window.DisplayPurAmount = function (kind, itemCode, classCode, entryNo, objText) {
		return displayAmountCommon(kind, itemCode, classCode, entryNo, objText, true);
	};

	window.setTaxPercentage = function (catCode, taxCode, objText) {
		var taxRoot = taxDetailsForPurchaseType(currentPurchaseType());
		if (!isNumeric(objText.value)) {
			alert("Enter Numeric Value");
			if (objText.select) {
				objText.select();
			}
			return false;
		}
		taxNodes(taxRoot).forEach(function (tax) {
			if (attr(tax, "CatCode") === trim(catCode) && attr(tax, "TaxCode") === trim(taxCode)) {
				setAttr(tax, "TaxValue", objText.value);
			}
		});
		window.popTax();
		window.CalculateItemValue(currentPurchaseType());
		return false;
	};

	window.setTaxAmount = function (catCode, taxCode, objText) {
		var taxRoot = taxDetailsForPurchaseType(currentPurchaseType());
		if (!isNumeric(objText.value)) {
			alert("Enter Numeric Value");
			if (objText.select) {
				objText.select();
			}
			return false;
		}
		taxNodes(taxRoot).forEach(function (tax) {
			if (attr(tax, "CatCode") === trim(catCode) && attr(tax, "TaxCode") === trim(taxCode)) {
				setAttr(tax, "TaxValue", objText.value);
				setAttr(tax, "TaxAmount", objText.value);
			}
		});
		window.popTax();
		window.CalculateItemValue(currentPurchaseType());
		return false;
	};

	window.CalculateTax = function (formula, basicValue, netValue, percentage, purchaseType) {
		return calculateFormulaTax(taxDetailsForPurchaseType(purchaseType), formula, basicValue, netValue, percentage, false);
	};

	window.popTax = function () {
		var taxRoot = taxDetailsForPurchaseType(currentPurchaseType());
		var basicTotal = toNumber(fieldValue("txtBasicValue"));
		var total = toNumber(fieldValue("txtAmount"));
		var invoiceAmount = total;
		var displayTotal = total;
		if (!taxRoot) {
			setRoundOffControls(invoiceAmount);
			updateHeaderTotals();
			return false;
		}
		taxNodes(taxRoot).forEach(function (tax) {
			var catCode = attr(tax, "CatCode");
			var taxCode = attr(tax, "TaxCode");
			var taxMode = attr(tax, "TaxMode");
			var taxValue = attr(tax, "TaxValue");
			var taxAmount = 0;
			if (catCode !== "0" || taxCode !== "0") {
				if (taxMode === "P") {
					taxAmount = window.CalculateTax(attr(tax, "TaxFormula"), basicTotal, total, taxValue, currentPurchaseType());
					setFieldValue("txtTaxPer" + catCode + taxCode, taxValue);
				} else if (taxMode === "K") {
					taxAmount = taxPackValue(tax);
				} else {
					taxAmount = toNumber(taxValue);
				}
				if (trim(attr(tax, "Rndoff")) === "1") {
					taxAmount = roundInvoice(taxAmount);
				}
				taxAmount = roundTo(taxAmount, 2);
				setAttr(tax, "TaxAmount", taxAmount);
				invoiceAmount += taxAmount;
				displayTotal += taxAmount;
				setFieldValue("txtTaxValue" + catCode + taxCode, formatNumber(taxAmount, 2));
				setFieldValue("txtSubTaxValue" + catCode + taxCode, formatNumber(displayTotal, 2));
			}
		});
		setRoundOffControls(invoiceAmount);
		updateTaxRoundOffNode(taxRoot);
		setAttr(taxRoot, "Basicvalue", formatNumber(basicTotal, 2));
		setAttr(taxRoot, "NettValue", formatNumber(total, 2));
		setAttr(taxRoot, "TotalTax", formatNumber(invoiceAmount - total, 2));
		setAttr(taxRoot, "SubTotal", formatNumber(invoiceAmount, 2));
		updateHeaderTotals();
		return false;
	};

	window.popTax1 = function () {
		var totalInvoice = 0;
		taxDetailsNodes().forEach(function (taxRoot) {
			var purchaseType = attr(taxRoot, "PurchaseType");
			var basicTotal = toNumber(fieldValue("txtBasicValue" + purchaseType));
			var total = toNumber(fieldValue("txtAmount" + purchaseType));
			var invoiceAmount = total;
			var taxDisplay = 0;
			var displayTax = 0;
			taxNodes(taxRoot).forEach(function (tax) {
				var catCode = attr(tax, "CatCode");
				var taxCode = attr(tax, "TaxCode");
				var taxMode = attr(tax, "TaxMode");
				var taxValue = attr(tax, "TaxValue");
				var taxAmount = 0;
				if (catCode !== "0" || taxCode !== "0") {
					if (taxMode === "P") {
						taxAmount = window.CalculateTax(attr(tax, "TaxFormula"), basicTotal, total, taxValue, purchaseType);
					} else if (taxMode === "K") {
						taxAmount = taxPackValue(tax);
					} else {
						taxAmount = toNumber(taxValue);
					}
					if (trim(attr(tax, "Rndoff")) === "1") {
						taxAmount = roundInvoice(taxAmount);
					}
					taxAmount = roundTo(taxAmount, 2);
					setAttr(tax, "TaxAmount", taxAmount);
					invoiceAmount += taxAmount;
					displayTax += taxAmount;
					if (trim(attr(tax, "AccHead")) !== "0") {
						taxDisplay += taxAmount;
					}
				}
			});
			updateTaxRoundOffNode(taxRoot);
			setAttr(taxRoot, "Basicvalue", formatNumber(basicTotal, 2));
			setAttr(taxRoot, "NettValue", formatNumber(total, 2));
			setAttr(taxRoot, "TotalTax", formatNumber(displayTax, 2));
			setAttr(taxRoot, "SubTotal", formatNumber(invoiceAmount, 2));
			setFieldValue("txtTot" + purchaseType, formatNumber(displayTax, 2));
			setFieldValue("txtTaxDisplay" + purchaseType, formatNumber(taxDisplay, 2));
			setFieldValue("txtSubTot" + purchaseType, formatNumber(invoiceAmount, 2));
			totalInvoice += invoiceAmount;
		});
		setRoundOffControls(totalInvoice);
		if (field("mTxtTotal")) {
			field("mTxtTotal").value = fieldValue("txtInvValue");
		}
		updateHeaderTotals();
		return false;
	};

	window.getTaxDet = function () {
		var root = invoiceRoot();
		var purchaseType = currentPurchaseType();
		var orgId = fieldValue("hOrgID");
		var response;
		if (!root) {
			return false;
		}
		taxDetailsNodes().forEach(function (node) {
			root.removeChild(node);
		});
		setDisabled("imgPurchaseDet", purchaseType !== "0");
		response = syncGet("XMLGetTaxDetails.asp?PurType=" + encodeURIComponent(purchaseType) + "&ForUnit=" + encodeURIComponent(orgId));
		if (response.responseText) {
			loadXmlIntoIsland("TaxFormData", response.responseText);
			children(xmlRoot(xmlIsland("TaxFormData"))).forEach(function (node) {
				root.appendChild(importNode(invoiceDoc(), node));
			});
		}
		if (headerNode()) {
			setAttr(headerNode(), "PurchaseType", purchaseType);
		}
		items().forEach(function (item) {
			setAttr(item, "PurchaseType", purchaseType);
		});
		taxDetailsNodes().forEach(function (taxRoot) {
			taxNodes(taxRoot).forEach(function (tax) {
				setAttr(tax, "ItemValue", "0");
			});
		});
		saveInvoice("InvItemValue");
		renderCurrentMode();
		return false;
	};

	window.showTaxFormPopUp = function (orgId, invoiceDate) {
		if (selectedValue("cmbPurType") === "0" || selectedIndex("cmbPurType") === 0) {
			alert("Select Purchase Type");
			if (field("cmbPurType")) {
				field("cmbPurType").focus();
			}
			return false;
		}
		openModal("invPurTaxFormPop.asp?OrgID=" + encodeURIComponent(orgId) + "&PurType=" + encodeURIComponent(selectedIndex("cmbPurType")) + "&PurTypeName=" + encodeURIComponent(selectedText("cmbPurType")) + "&Party=" + encodeURIComponent(fieldValue("txtPartyName")) + "&InvDt='" + encodeURIComponent(invoiceDate) + "'", "", "dialogHeight:500px;dialogWidth:900px;status:no", function (result) {
			var rootTax = xmlRoot(xmlIsland("TaxFormData"));
			children(xmlRoot(result)).forEach(function (node) {
				rootTax.appendChild(importNode(xmlDocument(xmlIsland("TaxFormData")), node));
			});
		});
		return false;
	};

	window.ShowPurchaseDet = function () {
		saveInvoice("InvItemValue");
		openModal("popItemPurType.asp", xmlIsland("InvoiceDet"), "dialogHeight:400px;dialogWidth:650px;status:no;help:no", function (result) {
			if (xmlRoot(result)) {
				replaceInvoiceRoot(result);
			}
			window.PurchaseTypeWise();
			window.popTax1();
		});
		return false;
	};

	window.ShowTaxDet = function (purchaseType) {
		openModal("popTaxDetails.asp?PurType=" + encodeURIComponent(purchaseType), xmlIsland("InvoiceDet"), "dialogLeft:0px;dialogTop:0Px;dialogHeight:600px;dialogWidth:500px;status:no", function (result) {
			if (xmlRoot(result)) {
				replaceInvoiceRoot(result);
			}
			window.popTax1();
			window.CalculateItemValue(purchaseType);
		});
		return false;
	};

	window.getReferenceDetail = function () {
		var receipt = document.getElementById("SpnRcptCode");
		window.bFlag = false;
		window.sInvAgainst = "";
		window.sInvRefNum = "";
		if (receipt && trim(receipt.textContent)) {
			window.bFlag = true;
			window.sInvAgainst = "Receipt";
			window.sInvRefNum = fieldValue("hRcptno");
		}
		return false;
	};

	window.displayItemDetail = function (orgId, flag) {
		var response;
		var root;
		window.getReferenceDetail();
		if (!window.bFlag) {
			alert("Select Reference Number");
			return false;
		}
		response = syncGet("XMLGetInvItem.asp?Flag=" + encodeURIComponent(flag) + "&sOrgID=" + encodeURIComponent(orgId) + "&InvoiceAgainst=" + encodeURIComponent(window.sInvAgainst) + "&RefNum=" + encodeURIComponent(window.sInvRefNum));
		if (response.responseText) {
			loadXmlIntoIsland("OutData", response.responseText);
			root = xmlRoot(xmlIsland("OutData"));
			children(root, "PartyType").forEach(function (partyTypeNode) {
				var select = field("cmbPartyType");
				if (select) {
					select.length = 0;
					children(partyTypeNode).forEach(function (node) {
						select.options[select.options.length] = new Option(attr(node, "SubTypeName"), attr(node, "SubTypeCode") + "|" + attr(node, "Type"));
					});
				}
			});
			children(root, "Header").forEach(function (header) {
				setFieldValue("hPartyCode", header.attributes.item(3) ? header.attributes.item(3).value : attr(header, "PartyCode"));
				setFieldValue("txtPartyName", header.attributes.item(4) ? header.attributes.item(4).value : attr(header, "PartyName"));
				if (field("cmbPartyType") && header.attributes.item(5) && header.attributes.item(6)) {
					field("cmbPartyType").value = header.attributes.item(5).value + "|" + header.attributes.item(6).value;
				}
				if (header.attributes.item(7) && trim(header.attributes.item(7).value) && trim(header.attributes.item(7).value) !== "0") {
					setFieldValue("txtSuppInvNo", header.attributes.item(7).value);
					setFieldValue("txtSuppInvDt", header.attributes.item(8) ? header.attributes.item(8).value : "");
				}
			});
		}
		return false;
	};

	function distributeItemValues(purchaseType) {
		var groupItems = itemsForPurchaseType(purchaseType);
		var taxRoot = taxDetailsForPurchaseType(purchaseType);
		var totalNet = toNumber(fieldValue("txtAmount" + (currentPurchaseType() === "0" ? purchaseType : "")));
		var totalItemValue = 0;
		if (!taxRoot) {
			groupItems.forEach(function (item) {
				setAttr(item, "Amount", attr(item, "NettBasic"));
			});
			return;
		}
		groupItems.forEach(function (item) {
			var qty = toNumber(attr(item, "Qty"));
			var itemValue = field("chkConsolidated") && field("chkConsolidated").checked ? toNumber(fieldValue("txtAmount")) : toNumber(attr(item, "NettBasic"));
			taxNodes(taxRoot).forEach(function (tax) {
				var taxAmount = 0;
				if (attr(tax, "CatCode") === "0" && attr(tax, "TaxCode") === "0") {
					return;
				}
				if (attr(tax, "TaxMode") === "P") {
					taxAmount = calculateFormulaTax(taxRoot, attr(tax, "TaxFormula"), itemValue, itemValue, attr(tax, "TaxValue"), true);
				} else if (totalNet > 0) {
					taxAmount = toNumber(attr(tax, "TaxAmount")) * itemValue / totalNet;
				}
				setAttr(tax, "ItemValue", Math.trunc(taxAmount * 100) / 100);
				if (toNumber(attr(tax, "AccHead")) === 0) {
					itemValue += taxAmount;
				}
			});
			setAttr(item, "Amount", itemValue);
			if (qty > 0) {
				setAttr(item, "ItemRate", roundTo(itemValue / qty, 4));
			}
			setAttr(item, "ItemValue", formatNumber(itemValue, 2));
			if (itemField("txtItemRateDisplay", attr(item, "ClassificationCode"), attr(item, "ItemCode"), attr(item, "SourceEntryNo"))) {
				itemField("txtItemRateDisplay", attr(item, "ClassificationCode"), attr(item, "ItemCode"), attr(item, "SourceEntryNo")).value = formatNumber(itemValue, 2);
			}
			totalItemValue += itemValue;
		});
		setFieldValue("txtTotalItemRate" + purchaseType, formatNumber(totalItemValue, 2));
	}

	window.CalcItemValueForEachItem = function (itemCode, classCode) {
		var item = items().filter(function (node) {
			return attr(node, "ItemCode") === trim(itemCode) && attr(node, "ClassificationCode") === trim(classCode);
		})[0];
		if (item) {
			window.CalculateItemValue(attr(item, "PurchaseType"));
		}
		return false;
	};

	window.CalculateItemValue = function (purchaseType) {
		if (field("chkConsolidated") && field("chkConsolidated").checked) {
			items().forEach(function (item) {
				setAttr(item, "Rate", fieldValue("txtAmount"));
				setAttr(item, "RatePerQtyUoM", fieldValue("txtAmount"));
				setAttr(item, "Amount", fieldValue("txtAmount"));
				setAttr(item, "ItemValue", fieldValue("txtAmount"));
				setAttr(item, "ItemRate", fieldValue("txtAmount"));
				setAttr(item, "NettBasic", fieldValue("txtAmount"));
			});
			window.popTax();
		}
		if (trim(purchaseType) && trim(purchaseType) !== "0") {
			distributeItemValues(purchaseType);
		} else {
			taxDetailsNodes().forEach(function (taxRoot) {
				distributeItemValues(attr(taxRoot, "PurchaseType"));
			});
			if (field("mTxtTotal")) {
				var total = 0;
				items().forEach(function (item) {
					total += toNumber(attr(item, "Amount"));
				});
				taxDetailsNodes().forEach(function (taxRoot) {
					total += toNumber(fieldValue("txtTaxDisplay" + attr(taxRoot, "PurchaseType")));
				});
				total += toNumber(fieldValue("txtRoundOff"));
				field("mTxtTotal").value = formatNumber(total, 2);
			}
		}
		return false;
	};

	window.CalcAmountForLastItem = function () {
		var list = items();
		var last = list[list.length - 1];
		var roundOffAccount = trim(fieldValue("hRoundOffAccHead"));
		var radio = controls("rdRndOff");
		var roundOffValue = toNumber(fieldValue("txtRoundOff"));
		if (last && roundOffAccount === "0" && radio && radio.length && radio[0].checked) {
			setAttr(last, "Amount", toNumber(attr(last, "Amount")) + roundOffValue);
		}
		list.forEach(function (item) {
			var classCode = attr(item, "ClassificationCode");
			var itemCode = attr(item, "ItemCode");
			var entryNo = attr(item, "SourceEntryNo");
			var qty = toNumber(attr(item, "Qty"));
			var amount = toNumber(attr(item, "Amount"));
			var vat = vatField(classCode, itemCode, entryNo);
			setAttr(item, "VAT", vat && vat.checked ? "Y" : "N");
			if (qty > 0 && amount > 0) {
				setAttr(item, "ItemRate", roundTo(amount / qty, 4));
				setAttr(item, "ItemValue", formatNumber(amount, 2));
			}
		});
		return false;
	};

	window.ItemValueCalculateTax = function (taxRoot, formula, basicValue, netValue, percentage) {
		return calculateFormulaTax(taxRoot, formula, basicValue, netValue, percentage, true);
	};

	window.AssignRoundOffValue = function () {
		var header = headerNode();
		var existingRoundOff = toNumber(attr(header, "RoundOff"));
		var newRoundOff = toNumber(fieldValue("txtRoundOff"));
		if (header && newRoundOff !== 0 && newRoundOff !== existingRoundOff) {
			setAttr(header, "InvValue", toNumber(attr(header, "InvValue")) - existingRoundOff + newRoundOff);
			setAttr(header, "RoundOff", newRoundOff);
			window.RoundOffInv();
		}
		return false;
	};

	window.RoundOffInv = function () {
		var header = headerNode();
		var radio = controls("rdRndOff");
		var invoiceValue = toNumber(attr(header, "InvValue"));
		var roundOffValue = toNumber(attr(header, "RoundOff"));
		if (!(radio && radio.length && radio[0].checked)) {
			invoiceValue -= roundOffValue;
			roundOffValue = 0;
		}
		setFieldValue("txtRoundOff", formatNumber(roundOffValue, 2));
		setFieldValue("txtInvValue", formatNumber(invoiceValue, 2));
		if (field("mTxtTotal")) {
			field("mTxtTotal").value = formatNumber(invoiceValue, 2);
		}
		return false;
	};

	window.AddRoundOffNode = function (roundOffAmount, invoiceValue) {
		var doc = invoiceDoc();
		if (headerNode()) {
			setAttr(headerNode(), "InvValue", invoiceValue);
		}
		taxDetailsNodes().forEach(function (taxRoot) {
			var roundNode = null;
			taxNodes(taxRoot).forEach(function (tax) {
				if (attr(tax, "CatCode") === "0" && attr(tax, "TaxCode") === "0") {
					roundNode = tax;
				}
			});
			if (!roundNode && doc) {
				roundNode = doc.createElement("Tax");
				taxRoot.appendChild(roundNode);
			}
			setAttr(roundNode, "CatCode", "0");
			setAttr(roundNode, "TaxCode", "0");
			setAttr(roundNode, "TaxMode", "0");
			setAttr(roundNode, "TaxFormula", "0");
			setAttr(roundNode, "TaxValue", roundOffAmount);
			setAttr(roundNode, "TaxAmount", roundOffAmount);
			setAttr(roundNode, "AccHead", fieldValue("hRoundOffAccHead"));
			setAttr(roundNode, "Formnumber", "0");
			setAttr(roundNode, "TransAmt", "0");
			setAttr(roundNode, "ItemValue", "0");
			roundNode.textContent = "ROUND OFF";
		});
		return false;
	};

	window.ShowPrefDet = function () {
		var headerParent = firstChild(invoiceRoot(), "InvoiceHeader");
		var header = headerNode();
		if (!header) {
			return false;
		}
		openModal("InvPurInvoiceEntryPref.asp?Mod1=" + encodeURIComponent(attr(header, "DespatchMode")) +
			"&Mop=" + encodeURIComponent(attr(header, "PaymentMode")) +
			"&IssueBank=" + encodeURIComponent(attr(header, "IssueBank")) +
			"&PayTerm=" + encodeURIComponent(attr(header, "PayTerms")) +
			"&Bop=" + encodeURIComponent(attr(header, "PricingBasis")) +
			"&Transporter=" + encodeURIComponent(attr(header, "Transporter")), headerParent, "dialogLeft:0px;dialogTop:0Px;dialogHeight:250px;dialogWidth:600px;status:no", function (result) {
				if (xmlRoot(result) && headerParent && headerParent.parentNode) {
					headerParent.parentNode.replaceChild(importNode(invoiceDoc(), xmlRoot(result)), headerParent);
				}
			});
		return false;
	};

	window.DeleteItems = function () {
		items().slice().forEach(function (item) {
			var checkbox = deleteField(attr(item, "ClassificationCode"), attr(item, "ItemCode"), attr(item, "SourceEntryNo"));
			if (checkbox && checkbox.checked) {
				window.DeleteItem(checkbox);
			}
		});
		return false;
	};

	window.DeleteItem = function (source) {
		var parts = String(source && source.name || "").split("A");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var item = matchingItem(itemCode, classCode, entryNo);
		var purchaseType = item ? attr(item, "PurchaseType") : "";
		var parent = itemDetailsNode();
		if (item && parent) {
			parent.removeChild(item);
		}
		if (purchaseType && itemsForPurchaseType(purchaseType).length === 0) {
			taxDetailsNodes().forEach(function (taxRoot) {
				if (attr(taxRoot, "PurchaseType") === purchaseType) {
					taxRoot.parentNode.removeChild(taxRoot);
				}
			});
		}
		renderCurrentMode();
		if (currentPurchaseType() === "0") {
			window.setValues1();
		} else {
			window.setValues();
		}
		return false;
	};

	window.setValues1 = function () {
		taxDetailsNodes().forEach(function (taxRoot) {
			var purchaseType = attr(taxRoot, "PurchaseType");
			updateTotalsForItems(itemsForPurchaseType(purchaseType), purchaseType);
		});
		window.popTax1();
		return false;
	};

	window.setValues = function () {
		updateTotalsForItems(items(), "");
		window.popTax();
		return false;
	};

	window.Next_Click = function (mode) {
		var header = headerNode();
		var amountControl = field("txtAmount") || field("mTxtTotal") || field("txtInvValue");
		var narration = field("txtNarration");
		var root = invoiceRoot();
		var doc = invoiceDoc();
		var existingNarration;
		var narrationNode;
		var consolidate = field("chkConsolidated") && field("chkConsolidated").checked ? "Y" : "N";
		if (!trim(fieldValue("txtSuppInvNo"))) {
			alert("Enter Supplier Invoice No.");
			if (field("txtSuppInvNo")) {
				field("txtSuppInvNo").focus();
			}
			return false;
		}
		if (dateDiffDays(controlDate("ctlDate"), fieldValue("hCurrDate")) < 0) {
			alert("Invoice date cannot be greater than the current date");
			return false;
		}
		if (selectedIndex("cmbPurType") === 0) {
			alert("Select Purchase Type ");
			if (field("cmbPurType")) {
				field("cmbPurType").focus();
			}
			return false;
		}
		if (selectedIndex("cmbInvCat") === 0) {
			alert("Select Category ");
			if (field("cmbInvCat")) {
				field("cmbInvCat").focus();
			}
			return false;
		}
		if (!narration || !trim(narration.value)) {
			alert("Enter Narration");
			if (narration) {
				narration.focus();
			}
			return false;
		}
		if (!amountControl || toNumber(amountControl.value) <= 0) {
			alert("Total Value Greater than zero");
			return false;
		}
		if (toNumber(fieldValue("txtInvValue")) <= 0) {
			alert("Invoice Value should be greater than 0");
			return false;
		}
		if (header) {
			setAttr(header, "InvValue", fieldValue("txtInvValue"));
			setAttr(header, "RoundOff", fieldValue("txtRoundOff"));
			setAttr(header, "Remarks", field("mTextAreaRemarks") ? field("mTextAreaRemarks").value : "");
			setAttr(header, "SuppInvNo", fieldValue("txtSuppInvNo"));
			setAttr(header, "SuppInvDt", controlDate("ctlDate"));
			setAttr(header, "ItemType", fieldValue("hItemType") || fieldValue("hItemtype"));
			setAttr(header, "InvCategory", selectedValue("cmbInvCat"));
			setAttr(header, "BillType", "P");
			setAttr(header, "GPNo", fieldValue("hGPNo"));
		}
		children(root, "Narration").forEach(function (node) {
			existingNarration = node;
		});
		if (existingNarration) {
			root.removeChild(existingNarration);
		}
		narrationNode = doc.createElement("Narration");
		narrationNode.textContent = narration.value;
		root.appendChild(narrationNode);
		if (mode === "V") {
			return true;
		}
		window.CalcAmountForLastItem();
		window.AddRoundOffNode(fieldValue("txtRoundOff"), fieldValue("txtInvValue"));
		saveInvoice("NewInvItemValue");
		form().action = "InvPurInvEntryService_AccDetails.asp?Consolidate=" + encodeURIComponent(consolidate);
		form().submit();
		return false;
	};

	window.ViewItemValues = function () {
		if (window.Next_Click("V")) {
			openModal("InvPurItemValueViewPop.asp", xmlIsland("InvoiceDet"), "dialogHeight:435px;dialogWidth:700px;Status:no;help:no", function () {});
		}
		return false;
	};

	window.SelectAll = function () {
		var checked;
		if (selectedIndex("cmbPurType") === 0) {
			alert("Select Purchase Type ");
			if (field("cmbPurType")) {
				field("cmbPurType").focus();
			}
			return false;
		}
		checked = field("ChkAll") && field("ChkAll").checked;
		items().forEach(function (item) {
			var vat = vatField(attr(item, "ClassificationCode"), attr(item, "ItemCode"), attr(item, "SourceEntryNo"));
			if (vat) {
				vat.checked = checked;
			}
		});
		return false;
	};
}(window, document));
