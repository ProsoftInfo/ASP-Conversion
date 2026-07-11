(function () {
	"use strict";

	var TaxRoot = null;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var normalized = String(value == null ? "" : value).replace(/,/g, "");
		var parsed = parseFloat(normalized);
		return isNaN(parsed) ? 0 : parsed;
	}

	function isNumeric(value) {
		return trim(value) !== "" && !isNaN(Number(String(value).replace(/,/g, "")));
	}

	function formatNumber(value, decimals) {
		return toNumber(value).toFixed(decimals == null ? 2 : decimals);
	}

	function round(value, decimals) {
		var factor = Math.pow(10, decimals || 0);
		return Math.round(toNumber(value) * factor) / factor;
	}

	function formField(name) {
		var form = document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
		return form && (form.elements[name] || form[name]) || null;
	}

	function fieldValue(name, fallback) {
		var field = formField(name);
		return field ? field.value : fallback;
	}

	function setFieldValue(name, value) {
		var field = formField(name);
		if (field) {
			field.value = value;
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || null;
	}

	function elementHtml(id) {
		var element = byId(id);
		return element ? element.innerHTML : "";
	}

	function setElementHtml(id, value) {
		var element = byId(id);
		if (element) {
			element.innerHTML = value;
		}
	}

	function getXmlRoot() {
		var data;
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
		data = window.TaxData;
		return data && data.documentElement || null;
	}

	function childElements(node, nodeName) {
		var result = [];
		var name = nodeName && String(nodeName).toLowerCase();
		if (!node) {
			return result;
		}
		for (var i = 0; i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!name || String(node.childNodes[i].nodeName).toLowerCase() === name)) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function firstChildElement(node, nodeName) {
		var matches = childElements(node, nodeName);
		return matches.length ? matches[0] : null;
	}

	function getDetailsRoot() {
		return firstChildElement(getXmlRoot(), "Details");
	}

	function getTaxRoot() {
		TaxRoot = firstChildElement(getXmlRoot(), "TaxDetails");
		return TaxRoot;
	}

	function attr(node, index) {
		return node && node.attributes ? node.attributes.item(index) : null;
	}

	function attrValue(node, index) {
		var attribute = attr(node, index);
		return attribute ? attribute.nodeValue : "";
	}

	function setAttrValue(node, index, value) {
		var attribute = attr(node, index);
		if (attribute) {
			attribute.nodeValue = value;
		}
	}

	function namedAttrValue(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setNamedAttrValue(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value);
		}
	}

	function eachTax(callback) {
		var root = getTaxRoot();
		var entries = childElements(root);
		for (var i = 0; i < entries.length; i += 1) {
			callback(entries[i]);
		}
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var nodes = [];
		if (!context) {
			return nodes;
		}
		if (context.selectNodes) {
			return context.selectNodes(expression);
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return nodes;
		}
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (var i = 0; i < found.snapshotLength; i += 1) {
			nodes.push(found.snapshotItem(i));
		}
		nodes.Item = function (index) {
			return this[index];
		};
		return nodes;
	}

	function nodeListLength(nodes) {
		return nodes ? nodes.length || 0 : 0;
	}

	function nodeListItem(nodes, index) {
		return nodes && (nodes.Item ? nodes.Item(index) : nodes[index]);
	}

	function findTaxNode(sCatCode, sTaxCode) {
		var found = null;
		eachTax(function (node) {
			if (found) {
				return;
			}
			if (attrValue(node, 0) === String(sCatCode) && attrValue(node, 1) === String(sTaxCode)) {
				found = node;
			}
		});
		return found;
	}

	window.setQty = function (objTemp, iSno, sCallTy) {
		var dQty;
		var dAmount = 0;
		var dTotal = 0;
		var dRate;
		var dRatePer = 1;
		var dDisPer = 0;
		var dDisAmount = 0;
		var dDisNewAmt;
		var dNewRate = fieldValue("txtRate" + iSno, 0);
		var dOldRate = elementHtml("tOldRate" + iSno);
		var detailRoot;
		var entries;

		if (trim(dNewRate) === "") {
			dNewRate = 0;
			setFieldValue("txtRate" + iSno, "0");
		}
		if (trim(dOldRate) === "") {
			dOldRate = 0;
			setElementHtml("tOldRate" + iSno, "0");
		}

		if (trim(objTemp.value) === "") {
			dQty = 0;
		} else if (!isNumeric(objTemp.value)) {
			alert("Enter Numeric values for Quantity");
			if (objTemp.focus) {
				objTemp.focus();
			}
			return;
		} else if (toNumber(objTemp.value) < 0 || toNumber(objTemp.value) > 9999999.999) {
			alert("Quantity should be >=0 and < 9999999.999");
			if (objTemp.focus) {
				objTemp.focus();
			}
			return;
		} else {
			dQty = objTemp.value;
		}

		detailRoot = getDetailsRoot();
		entries = childElements(detailRoot, "Entry");
		for (var i = 0; i < entries.length; i += 1) {
			if (namedAttrValue(entries[i], "No") === String(iSno)) {
				setNamedAttrValue(entries[i], "Qty", dQty);
				dRate = formatNumber(dNewRate, 2);
				setNamedAttrValue(entries[i], "Rate", dNewRate);
				if (toNumber(dRatePer) === 0) {
					dRatePer = 1;
				}
				dAmount = (toNumber(dRate) / toNumber(dRatePer)) * toNumber(dQty);
				dDisPer = fieldValue("hDisPer" + iSno, 0);
				if (toNumber(dDisPer) !== 0) {
					dDisAmount = (toNumber(trim(dDisPer)) / 100) * dAmount;
					dAmount = formatNumber(dAmount, 2);
					dDisAmount = formatNumber(dDisAmount, 2);
					dDisNewAmt = fieldValue("txtDis" + iSno, 0);
					if (toNumber(dRate) > 0) {
						if (String(dDisAmount) !== String(dDisNewAmt)) {
							dDisPer = formatNumber((toNumber(dDisNewAmt) / toNumber(dRate)) * 100, 2);
							setNamedAttrValue(entries[i], "DisPer", dDisPer);
							dDisAmount = dDisNewAmt;
						}
					} else {
						dDisAmount = 0;
					}
					if (toNumber(dQty) === 0) {
						dDisAmount = 0;
					}
					dAmount = toNumber(dAmount) - toNumber(dDisAmount);
				} else {
					dDisAmount = fieldValue("txtDis" + iSno, 0);
					if (toNumber(dRate) > 0) {
						dDisPer = formatNumber((toNumber(dDisAmount) / toNumber(dRate)) * 100, 2);
					} else {
						dDisAmount = 0;
					}
					if (toNumber(dQty) === 0) {
						dDisAmount = 0;
					}
					dAmount = dAmount - toNumber(dDisAmount);
					setNamedAttrValue(entries[i], "DisPer", dDisPer);
				}
				dTotal += toNumber(dAmount);
				setNamedAttrValue(entries[i], "DisAmount", dDisAmount);
				setNamedAttrValue(entries[i], "Amount", dAmount);
			} else {
				dTotal += toNumber(namedAttrValue(entries[i], "Amount"));
			}
		}

		setFieldValue("txtTotal", formatNumber(dTotal, 2));
		setFieldValue("txtDis" + iSno, formatNumber(dDisAmount, 2));
		setFieldValue("txtAmount" + iSno, formatNumber(dAmount, 2));
		setNamedAttrValue(detailRoot, "BasicValue", dTotal);
		setNamedAttrValue(detailRoot, "Discount", dDisPer);
		setNamedAttrValue(detailRoot, "ActualValue", dTotal);
		window.popTax();

		if (String(fieldValue("SelCrAgain", "")) === "A") {
			setFieldValue("txtCrNoteValue", fieldValue("txtInvValue", ""));
		}
		if (typeof window.SetRetVal === "function") {
			window.SetRetVal(formField("SelCrAgain"), "0");
		}
	};

	window.setTotal = function (objTemp, iSno) {
		var dAmount;
		var dTotal = 0;
		var detailRoot;
		var entries;

		if (trim(objTemp.value) === "") {
			dAmount = 0;
		} else if (!isNumeric(objTemp.value)) {
			alert("Enter Numeric values for Amount");
			if (objTemp.focus) {
				objTemp.focus();
			}
			return;
		} else if (toNumber(objTemp.value) < 0 || toNumber(objTemp.value) > 9999999999.99) {
			alert("Amount should be >=0 and < 9999999999.99");
			if (objTemp.focus) {
				objTemp.focus();
			}
			return;
		} else {
			dAmount = objTemp.value;
		}

		detailRoot = getDetailsRoot();
		entries = childElements(detailRoot);
		for (var i = 0; i < entries.length; i += 1) {
			if (namedAttrValue(entries[i], "No") === String(iSno)) {
				setNamedAttrValue(entries[i], "Amount", dAmount);
				dTotal += toNumber(dAmount);
			} else {
				dTotal += toNumber(namedAttrValue(entries[i], "Amount"));
			}
		}

		setFieldValue("txtTotal", formatNumber(dTotal, 2));
		objTemp.value = formatNumber(dAmount, 2);
		setNamedAttrValue(detailRoot, "BasicValue", dTotal);
		setNamedAttrValue(detailRoot, "Discount", "0");
		setNamedAttrValue(detailRoot, "ActualValue", dTotal);
		window.popTax();
		setFieldValue("txtCrNoteValue", fieldValue("txtInvValue", ""));
	};

	window.setTaxPercentage = function (sCatCode, sTaxCode, objText) {
		var value = trim(objText && objText.value);
		var taxNode = findTaxNode(sCatCode, sTaxCode);
		if (value !== "" && !isNumeric(value)) {
			alert("Enter Numeric Value");
			if (objText && objText.select) {
				objText.select();
			}
			return;
		}
		if (value !== "" && taxNode) {
			setAttrValue(taxNode, 4, objText.value);
			window.popTax();
		}
		if (String(fieldValue("SelCrAgain", "")) === "A") {
			setFieldValue("txtCrNoteValue", fieldValue("txtInvValue", ""));
		}
	};

	window.setTaxAmount = function (sCatCode, sTaxCode, objText) {
		var value = trim(objText && objText.value);
		var dInvAmount = fieldValue("txtTotal", 0);
		var taxNode = findTaxNode(sCatCode, sTaxCode);
		var sTaxMode;
		var sFormula;
		if (value !== "" && !isNumeric(value)) {
			alert("Enter Numeric Value");
			if (objText && objText.select) {
				objText.select();
			}
			return;
		}
		if (value !== "" && taxNode) {
			sTaxMode = attrValue(taxNode, 2);
			sFormula = attrValue(taxNode, 3);
			if (sTaxMode === "P") {
				setAttrValue(taxNode, 4, window.calPercentage(sFormula, dInvAmount, dInvAmount, objText.value));
			} else {
				setAttrValue(taxNode, 4, objText.value);
				setAttrValue(taxNode, 5, objText.value);
			}
			window.popTax();
		}
		if (String(fieldValue("SelCrAgain", "")) === "A") {
			setFieldValue("txtCrNoteValue", fieldValue("txtInvValue", ""));
		}
	};

	window.popTax = function () {
		var rootNode = getXmlRoot();
		var root = getTaxRoot();
		var dInvAmount;
		var dBasicTotal;
		var dTotal;
		var roundedValue;
		var roundedOff;
		var roundOffNodes;
		var taxNodes;
		var roundOffTax;
		if (!root) {
			return;
		}
		dInvAmount = toNumber(fieldValue("txtTotal", 0));
		dBasicTotal = dInvAmount;
		dTotal = dInvAmount;

		eachTax(function (node) {
			var sCatCode = attrValue(node, 0);
			var sTaxCode = attrValue(node, 1);
			var sTaxMode = attrValue(node, 2);
			var sFormula = attrValue(node, 3);
			var dTaxValue = attrValue(node, 4);
			var sTotpackvalue = attrValue(node, 8);
			var iRndOff = sCatCode === "0" && sTaxCode === "0" && sTaxMode === "0" ? 0 : toNumber(attrValue(node, 7));
			var dTax;
			if (!formField("txtTaxValue" + sCatCode + sTaxCode)) {
				return;
			}
			if (sTaxMode === "P") {
				dTax = window.CalculateTax(sFormula, dBasicTotal, dTotal, dTaxValue);
				setFieldValue("txtTaxPer" + sCatCode + sTaxCode, dTaxValue);
			} else if (sTaxMode === "Q") {
				dTax = toNumber(dTaxValue) * toNumber(window.Qtysum());
			} else if (sTaxMode === "K") {
				dTax = toNumber(sTotpackvalue) * toNumber(window.Qtysum());
			} else {
				dTax = dTaxValue;
			}

			if (!(sCatCode === "0" && sTaxCode === "0" && sTaxMode === "0") && iRndOff === 1) {
				dTax = round(dTax, 0);
			}
			if (String(fieldValue("SelCrAgain", "")) === "A" && sCatCode !== "0" && sTaxCode !== "0" && sTaxMode !== "0") {
				dTax = 0;
			}

			dInvAmount += toNumber(dTax);
			dTax = formatNumber(dTax, 2);
			setAttrValue(node, 5, dTax);
			setFieldValue("txtTaxValue" + sCatCode + sTaxCode, dTax);
		});

		dInvAmount = formatNumber(dInvAmount, 2);
		setFieldValue("txtInvValue", formatNumber(round(dInvAmount, 0), 2));
		roundedValue = fieldValue("txtInvValue", 0);
		roundedOff = round(toNumber(roundedValue) - toNumber(dInvAmount), 2);
		roundOffNodes = selectNodes(root, "//TaxDetails[@RoundOffValue]");

		setAttrValue(root, 0, fieldValue("txtInvValue", ""));
		setAttrValue(root, 1, fieldValue("txtTotal", ""));
		setAttrValue(root, 2, fieldValue("txtTotal", ""));
		if (nodeListLength(roundOffNodes) !== 0) {
			setAttrValue(root, 3, roundedOff);
		}

		taxNodes = selectNodes(rootNode, "//Tax[@CatCode=0 and @TaxCode=0 and @TaxMode=0]");
		if (nodeListLength(taxNodes) !== 0) {
			roundOffTax = nodeListItem(taxNodes, 0);
			setAttrValue(roundOffTax, 5, roundedOff);
			setFieldValue("txtTaxValue00", roundedOff);
		}
	};

	window.CalculateTax = function (sFormula, dBValue, dDValue, dPercentage) {
		var parts = String(sFormula || "").split(",");
		var dTaxAmount;
		var startIndex;
		var token;
		if (trim(parts[0]) === "BV") {
			dTaxAmount = toNumber(dBValue);
			startIndex = 1;
		} else if (trim(parts[0]) === "BD") {
			dTaxAmount = toNumber(dDValue);
			startIndex = 1;
		} else {
			dTaxAmount = 0;
			startIndex = 0;
		}
		for (var i = startIndex; i < parts.length; i += 1) {
			token = trim(parts[i]).split("#");
			eachTax(function (node) {
				if (attrValue(node, 0) === trim(token[0]) && attrValue(node, 1) === trim(token[1])) {
					dTaxAmount += toNumber(attrValue(node, 5));
				}
			});
		}
		return formatNumber(dTaxAmount * (toNumber(dPercentage) / 100), 2);
	};

	window.calPercentage = function (sFormula, dBValue, dDValue, dAmount) {
		var parts = String(sFormula || "").split(",");
		var dTaxAmount;
		var startIndex;
		var token;
		if (trim(sFormula) === "" || trim(sFormula) === "0") {
			return;
		}
		if (trim(parts[0]) === "BV") {
			dTaxAmount = toNumber(dBValue);
			startIndex = 1;
		} else if (trim(parts[0]) === "BD") {
			dTaxAmount = toNumber(dDValue);
			startIndex = 1;
		} else {
			dTaxAmount = 0;
			startIndex = 0;
		}
		for (var i = startIndex; i < parts.length; i += 1) {
			token = trim(parts[i]).split("#");
			eachTax(function (node) {
				if (attrValue(node, 0) === trim(token[0]) && attrValue(node, 1) === trim(token[1])) {
					dTaxAmount += toNumber(attrValue(node, 5));
				}
			});
		}
		return formatNumber(toNumber(dAmount) / dTaxAmount * 100, 2);
	};

	window.Qtysum = function () {
		var root = getXmlRoot();
		var nodes = selectNodes(root, "//Details/Entry");
		var qty = 0;
		for (var i = 0; i < nodeListLength(nodes); i += 1) {
			qty += toNumber(attrValue(nodeListItem(nodes, i), 3));
		}
		return qty;
	};
	window.QtySum = window.Qtysum;
}());
