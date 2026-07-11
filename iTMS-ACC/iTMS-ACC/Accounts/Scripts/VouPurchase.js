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

	function getTaxRoot() {
		var data;
		var root;
		var child;
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
		data = window.TaxData;
		root = data && data.documentElement;
		TaxRoot = null;
		if (!root) {
			return null;
		}
		for (var i = 0; i < root.childNodes.length; i += 1) {
			child = root.childNodes[i];
			if (child.nodeType === 1 && String(child.nodeName).toLowerCase() === "taxdetails") {
				TaxRoot = child;
				break;
			}
		}
		return TaxRoot;
	}

	function eachTax(callback) {
		var root = getTaxRoot();
		if (!root) {
			return;
		}
		for (var i = 0; i < root.childNodes.length; i += 1) {
			if (root.childNodes[i].nodeType === 1) {
				callback(root.childNodes[i]);
			}
		}
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

	window.setTaxPercentage = function (sCatCode, sTaxCode, objText) {
		var value = trim(objText && objText.value);
		var taxNode;
		getTaxRoot();
		if (toNumber(value) < 0) {
			alert("Tax Value Should be Greater than Zero ");
			if (objText && objText.focus) {
				objText.focus();
			}
			return;
		}
		if (value === "") {
			return;
		}
		if (!isNumeric(value)) {
			alert("Enter Numeric Value");
			if (objText && objText.select) {
				objText.select();
			}
			return;
		}
		taxNode = findTaxNode(sCatCode, sTaxCode);
		if (taxNode) {
			setAttrValue(taxNode, 4, objText.value);
			window.popTax();
		}
	};

	window.setTaxAmount = function (sCatCode, sTaxCode, objText) {
		var value = trim(objText && objText.value);
		var taxNode;
		var sTaxMode;
		var sFormula;
		var dBasicTotal = fieldValue("hBasicValue", 0);
		var dTotal = fieldValue("hAmount", 0);
		getTaxRoot();
		if (value === "") {
			return;
		}
		if (!isNumeric(value)) {
			alert("Enter Numeric Value");
			if (objText && objText.select) {
				objText.select();
			}
			return;
		}
		taxNode = findTaxNode(sCatCode, sTaxCode);
		if (taxNode) {
			sTaxMode = attrValue(taxNode, 2);
			sFormula = attrValue(taxNode, 3);
			if (sTaxMode === "P") {
				setAttrValue(taxNode, 4, window.calPercentage(sFormula, dBasicTotal, dTotal, objText.value));
			} else {
				setAttrValue(taxNode, 4, objText.value);
				setAttrValue(taxNode, 5, objText.value);
			}
			window.popTax();
		}
	};

	window.popTax = function () {
		var dInvAmount;
		var dBasicTotal;
		var dTotal;
		var dRoundedInvvalue;
		var dRoundedoff;
		var root = getTaxRoot();
		if (!root) {
			return;
		}
		dInvAmount = toNumber(fieldValue("hAmount", 0));
		dBasicTotal = fieldValue("hBasicValue", 0);
		dTotal = fieldValue("hAmount", 0);

		eachTax(function (node) {
			var sCatCode = attrValue(node, 0);
			var sTaxCode = attrValue(node, 1);
			var sTaxMode = attrValue(node, 2);
			var sFormula = attrValue(node, 3);
			var dTaxValue = formatNumber(attrValue(node, 4), 2);
			var iRndOff = sCatCode !== "0" && sTaxCode !== "0" ? toNumber(attrValue(node, 7)) : 0;
			var dTax = dTaxValue;
			if (sTaxMode === "P") {
				if (sCatCode !== "0" && sTaxCode !== "0") {
					dTax = window.CalculateTax(sFormula, dBasicTotal, dTotal, dTaxValue);
					setFieldValue("txtTaxPer" + sCatCode + sTaxCode, dTaxValue);
				}
			}
			if (iRndOff === 1) {
				dTax = round(dTax, 0);
			}
			dInvAmount += toNumber(dTax);
			dTax = formatNumber(dTax, 2);
			setAttrValue(node, 5, dTax);
			if (sCatCode !== "0" && sTaxCode !== "0") {
				setFieldValue("txtTaxValue" + sCatCode + sTaxCode, dTax);
			}
		});

		if (String(fieldValue("hRndChk", "")) === "1") {
			dRoundedInvvalue = round(dInvAmount, 0);
			dRoundedoff = round(toNumber(dRoundedInvvalue) - dInvAmount, 2);
		} else {
			dRoundedInvvalue = dInvAmount;
			dRoundedoff = 0;
		}
		dRoundedInvvalue = formatNumber(dRoundedInvvalue, 2);
		setFieldValue("txtroundoff", dRoundedoff);
		setFieldValue("txtInvValue", dRoundedInvvalue);
		setAttrValue(root, 0, dRoundedInvvalue);
		setAttrValue(root, 1, dTotal);
		setAttrValue(root, 2, dRoundedInvvalue);
		setAttrValue(root, 3, dRoundedoff);
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
		return dTaxAmount > 0 ? formatNumber(toNumber(dAmount) / dTaxAmount * 100, 2) : 0;
	};
}());
