(function (window, document) {
	"use strict";

	var itemTaxData = modalArgs();
	var purchType = "";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
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
		return frm && frm.elements[name];
	}

	function itemDocument() {
		if (itemTaxData && itemTaxData.XMLDocument) {
			return itemTaxData.XMLDocument;
		}
		if (itemTaxData && itemTaxData._doc) {
			return itemTaxData._doc;
		}
		if (itemTaxData && itemTaxData.nodeType === 9) {
			return itemTaxData;
		}
		if (itemTaxData && itemTaxData.ownerDocument) {
			return itemTaxData.ownerDocument;
		}
		return null;
	}

	function itemRoot() {
		if (itemTaxData && itemTaxData.documentElement) {
			return itemTaxData.documentElement;
		}
		if (itemTaxData && itemTaxData.nodeType === 1) {
			return itemTaxData;
		}
		return itemDocument() && itemDocument().documentElement;
	}

	function xpathLiteral(value) {
		var text = String(value == null ? "" : value);
		if (text.indexOf("'") === -1) {
			return "'" + text + "'";
		}
		return '"' + text.replace(/"/g, "") + '"';
	}

	function selectNodes(context, expression) {
		return context && typeof context.selectNodes === "function" ? context.selectNodes(expression) : [];
	}

	function childElements(node) {
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1;
		});
	}

	function attr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function setSpan(id, value) {
		var el = document.getElementById(id);
		if (el) {
			el.innerHTML = value == null ? "" : String(value);
		}
	}

	function spanNumber(id) {
		var el = document.getElementById(id);
		return numberValue(el && el.innerHTML);
	}

	function appendCell(row, text, className, align) {
		var cell = row.insertCell();
		cell.className = className || "ExcelSerial";
		cell.align = align || "left";
		cell.innerHTML = text == null ? "" : String(text);
		return cell;
	}

	function taxDetailsForPurchaseType(root, purchaseType) {
		var nodes = selectNodes(root, "//TaxDetails[@PurchaseType=" + xpathLiteral(purchaseType) + "]");
		return nodes.length ? nodes.item(0) : null;
	}

	function findTaxNode(taxRoot, catCode, taxCode) {
		var nodes = childElements(taxRoot);
		for (var index = 0; index < nodes.length; index += 1) {
			if (attr(nodes[index], "CatCode") === trim(catCode) && attr(nodes[index], "TaxCode") === trim(taxCode)) {
				return nodes[index];
			}
		}
		return null;
	}

	function qtySum() {
		return typeof window.Qtysum === "function" ? numberValue(window.Qtysum()) : 0;
	}

	function setReturnValue() {
		var root = itemRoot();
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	window.init = function () {
		var root = itemRoot();
		var purchaseName = trim(field("hPurTypeName") && field("hPurTypeName").value);
		var itemNodes;
		var taxNodes;
		var itemName = "";
		var netTotal = 0;
		var basicTotal = 0;
		var table = document.getElementById("tblTaxDet");
		var row;
		var taxValue = "";

		purchType = trim(field("hPurType") && field("hPurType").value);
		itemNodes = selectNodes(root, "//ItemDetails/Item[@PurchaseType=" + xpathLiteral(purchType) + "]");

		for (var index = 0; index < itemNodes.length; index += 1) {
			var item = itemNodes.item(index);
			var qty = numberValue(attr(item, "Quantity") || item.attributes.item(4) && item.attributes.item(4).nodeValue);
			var rate = numberValue(attr(item, "RatePerQtyUoM"));
			var discount = numberValue(item.attributes.item(7) && item.attributes.item(7).nodeValue);
			if (itemName) {
				itemName += ",";
			}
			itemName += attr(item, "ItmDescription");
			basicTotal += qty * rate;
			netTotal += (qty * rate) - discount;
		}

		setSpan("spnItemName", itemName);
		setSpan("spnPurType", purchaseName);
		setSpan("spnAmount", formatNumber(round(netTotal, 2), 2));
		setSpan("spnBasicValue", formatNumber(round(basicTotal, 2), 2));

		taxNodes = selectNodes(root, "//TaxDetails[@PurchaseType=" + xpathLiteral(purchType) + "]/Tax");
		for (index = 0; index < taxNodes.length; index += 1) {
			var tax = taxNodes.item(index);
			var taxName = tax.textContent || "";
			var catCode = attr(tax, "CatCode");
			var taxCode = attr(tax, "TaxCode");
			var taxMode = attr(tax, "TaxMode");
			var taxPercent = "";

			if (catCode === "0" || taxCode === "0") {
				continue;
			}
			if (taxMode === "F") {
				taxValue = attr(tax, "TaxValue");
			} else {
				taxPercent = attr(tax, "TaxValue");
				taxValue = "";
			}

			row = table.insertRow(table.rows.length);
			appendCell(row, index + 1, "ExcelSerial", "right");
			appendCell(row, taxName, "ExcelSerial", "left");

			var percentCell = row.insertCell();
			percentCell.align = "center";
			if (taxMode !== "F") {
				var percentInput = document.createElement("input");
				percentInput.type = "text";
				percentInput.name = "txtTaxPer" + catCode + taxCode;
				percentInput.value = taxPercent;
				percentInput.size = 4;
				percentInput.style.textAlign = "right";
				percentInput.className = "FormElem";
				percentInput.onblur = function (passCatCode, passTaxCode, input) {
					return function () {
						window.setTaxPercentage(passCatCode, passTaxCode, input);
					};
				}(catCode, taxCode, percentInput);
				percentCell.appendChild(percentInput);
				percentCell.appendChild(document.createTextNode(" %"));
				percentCell.className = "ExcelFieldCell";
			} else {
				percentCell.className = "ExcelSerial";
			}
			percentCell.width = "10";

			var valueCell = row.insertCell();
			var valueInput = document.createElement("input");
			valueInput.type = "text";
			valueInput.name = "txtTaxValue" + catCode + taxCode;
			valueInput.value = taxValue;
			valueInput.size = 11;
			valueInput.style.textAlign = "right";
			if (taxMode === "F") {
				valueInput.className = "FormElem";
				valueInput.onblur = function (passCatCode, passTaxCode, input) {
					return function () {
						window.setTaxAmount(passCatCode, passTaxCode, input);
					};
				}(catCode, taxCode, valueInput);
			} else {
				valueInput.readOnly = true;
				valueInput.className = "FormElemRead";
			}
			valueCell.className = "ExcelFieldCell";
			valueCell.align = "center";
			valueCell.width = "10";
			valueCell.appendChild(valueInput);
		}

		row = table.insertRow(table.rows.length);
		var totalLabel = appendCell(row, "Total", "ExcelSerial", "right");
		totalLabel.width = "10";
		totalLabel.colSpan = 3;
		var totalCell = row.insertCell();
		var totalInput = document.createElement("input");
		totalInput.type = "text";
		totalInput.name = "txtTotTax";
		totalInput.value = taxValue || "0";
		totalInput.size = 11;
		totalInput.readOnly = true;
		totalInput.style.textAlign = "right";
		totalInput.className = "FormElemRead";
		totalCell.appendChild(totalInput);
		totalCell.width = "10";
		totalCell.className = "ExcelFieldCell";
		totalCell.align = "center";

		window.popTax();
	};

	window.setTaxPercentage = function (catCode, taxCode, input) {
		var root = itemRoot();
		var taxRoot = taxDetailsForPurchaseType(root, purchType);
		var taxNode = findTaxNode(taxRoot, catCode, taxCode);
		if (!trim(input.value)) {
			return;
		}
		if (isNaN(Number(input.value))) {
			alert("Enter Numeric Value");
			if (input.select) {
				input.select();
			}
			return;
		}
		if (taxNode) {
			taxNode.setAttribute("TaxValue", input.value);
			window.popTax();
		}
	};

	window.popTax = function () {
		var root = itemRoot();
		var taxRoot = taxDetailsForPurchaseType(root, purchType);
		var basicTotal = spanNumber("spnBasicValue");
		var displayTotal = spanNumber("spnAmount");
		var totalTax = 0;
		if (!taxRoot) {
			return;
		}

		childElements(taxRoot).forEach(function (taxNode) {
			var catCode = attr(taxNode, "CatCode");
			var taxCode = attr(taxNode, "TaxCode");
			var taxMode = attr(taxNode, "TaxMode");
			var formula = attr(taxNode, "TaxFormula");
			var taxValue = attr(taxNode, "TaxValue");
			var taxAmount = 0;

			if (catCode === "0" || taxCode === "0") {
				return;
			}

			if (taxMode === "P") {
				taxAmount = window.CalculateTax(formula, basicTotal, displayTotal, taxValue, purchType);
				if (field("txtTaxPer" + catCode + taxCode)) {
					field("txtTaxPer" + catCode + taxCode).value = taxValue;
				}
			} else if (taxMode === "Q") {
				taxAmount = numberValue(taxValue) * qtySum();
			} else if (taxMode === "K") {
				var packageValue = 0;
				selectNodes(taxRoot, ".//Tax[@CatCode=" + xpathLiteral(catCode) + " and @TaxCode=" + xpathLiteral(taxCode) + "]/Taxpack").forEach(function (packNode) {
					packageValue += numberValue(packNode.attributes.item(3) && packNode.attributes.item(3).nodeValue);
				});
				taxAmount = packageValue;
			} else {
				taxAmount = numberValue(taxValue);
			}

			if (attr(taxNode, "Rndoff") === "1") {
				taxAmount = typeof window.RndOff === "function" ? window.RndOff(taxAmount) : Math.round(taxAmount);
			}

			taxAmount = formatNumber(taxAmount, 2);
			taxNode.setAttribute("TaxAmount", taxAmount);
			if (field("txtTaxValue" + catCode + taxCode)) {
				field("txtTaxValue" + catCode + taxCode).value = taxAmount;
			}
			totalTax += numberValue(taxAmount);
		});

		taxRoot.setAttribute("Basicvalue", formatNumber(basicTotal, 2));
		taxRoot.setAttribute("NettValue", formatNumber(displayTotal, 2));
		taxRoot.setAttribute("TotalTax", formatNumber(totalTax, 2));
		taxRoot.setAttribute("SubTotal", formatNumber(displayTotal + totalTax, 2));
		if (field("txtTotTax")) {
			field("txtTotTax").value = formatNumber(totalTax, 2);
		}
	};

	window.CalculateTax = function (formula, basicValue, displayValue, percentage, purchaseType) {
		var root = itemRoot();
		var taxRoot = purchaseType === "0" ? selectNodes(root, "//TaxDetails").item(0) : taxDetailsForPurchaseType(root, purchaseType);
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
			var tokens = trim(parts[index]).split("#");
			var node = findTaxNode(taxRoot, tokens[0], tokens[1]);
			if (node) {
				taxAmount += numberValue(attr(node, "TaxAmount"));
			}
		}

		if (trim(percentage)) {
			return taxAmount * (numberValue(percentage) / 100);
		}
		return taxAmount;
	};

	window.setTaxAmount = function (catCode, taxCode, input) {
		var root = itemRoot();
		var taxRoot = taxDetailsForPurchaseType(root, purchType);
		var taxNode = findTaxNode(taxRoot, catCode, taxCode);
		var taxMode;
		var formula;

		if (!trim(input.value)) {
			return;
		}
		if (isNaN(Number(input.value))) {
			alert("Enter Numeric Value");
			if (input.select) {
				input.select();
			}
			return;
		}
		if (!taxNode) {
			return;
		}

		taxMode = attr(taxNode, "TaxMode");
		formula = attr(taxNode, "TaxFormula");
		if (taxMode === "F") {
			taxNode.setAttribute("TaxValue", input.value);
			taxNode.setAttribute("TaxAmount", input.value);
		} else if (taxMode === "P") {
			taxNode.setAttribute("TaxValue", window.CalculateTax(formula, spanNumber("spnBasicValue"), spanNumber("spnAmount"), input.value, purchType));
		} else {
			taxNode.setAttribute("TaxValue", input.value);
			taxNode.setAttribute("TaxAmount", input.value);
		}
		window.popTax();
	};

	window.addEventListener("beforeunload", setReturnValue);
})(window, document);
