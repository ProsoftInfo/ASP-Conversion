(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var rootDoc = null;
	var lotNo = "";
	var issueDate = "";
	var qtyTotal = 0;

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
		return frm && frm.elements ? frm.elements[name] || null : null;
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

	function normalizeLot(value) {
		var lot = trim(value);
		return lot.toUpperCase() === "NULL" || lot.toUpperCase() === "N/A" ? "0" : lot;
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

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
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

	function attrAt(node, index) {
		var attr = node && node.attributes && node.attributes.item(index);
		return trim(attr ? attr.nodeValue || attr.value || "" : "");
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, trim(value));
		}
	}

	function fieldValue(name) {
		var item = field(name);
		return trim(item ? item.value : "");
	}

	function createElement(name) {
		if (rootDoc && rootDoc.createElement) {
			return rootDoc.createElement(name);
		}
		return document.implementation.createDocument("", name, null).documentElement;
	}

	function rowCount() {
		return toNumber(field("hiCtr") && field("hiCtr").value);
	}

	function currentLotNode() {
		var nodes = elementChildren(root, "LotDet");
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "LotNo") === lotNo && getAttr(nodes[i], "IssueDate") === issueDate) {
				return nodes[i];
			}
		}
		return null;
	}

	function removeSerialDetails(lotNode) {
		var details = elementChildren(lotNode, "SerialDetails");
		for (var i = 0; i < details.length; i += 1) {
			lotNode.removeChild(details[i]);
		}
	}

	function setReturnValue() {
		if (!root) {
			root = xmlRoot(objTemp);
		}
		if (root) {
			setAttr(root, "SerQtyRet", qtyTotal);
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	window.fnInit = function (arg) {
		var parts = String(arg || "").split("`");
		var lotNode;
		var details;
		var qty;

		objTemp = modalArgs();
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		if (!rootDoc && root) {
			rootDoc = root.ownerDocument;
		}

		lotNo = normalizeLot(parts[5] || "");
		issueDate = parts[9] || "";

		if (!rowCount() || !root) {
			return;
		}

		lotNode = currentLotNode();
		details = elementChildren(lotNode, "SerialDetails");
		for (var i = 0; i < details.length; i += 1) {
			qty = field("txtQty" + (i + 1));
			if (qty) {
				qty.value = getAttr(details[i], "Qty") || attrAt(details[i], 1);
			}
		}
	};

	window.CheckSubmit = function () {
		var rows = rowCount();
		var lotNode;
		var detail;
		var qty;
		var stock;

		qtyTotal = 0;
		if (!rows) {
			return false;
		}

		for (var i = 1; i <= rows; i += 1) {
			qty = field("txtQty" + i);
			stock = field("txtStQty" + i);
			if (!qty || trim(qty.value) === "") {
				alert("Enter Quantity");
				if (qty) {
					qty.select();
				}
				return false;
			}
			if (!isNumericText(qty.value)) {
				alert("Enter Numerals Only");
				qty.select();
				return false;
			}
			if (toNumber(qty.value) > toNumber(stock && stock.value)) {
				alert("Return Quantity should be equal to or less than Quantity Remaining");
				qty.select();
				return false;
			}
			qtyTotal += toNumber(qty.value);
		}

		lotNode = currentLotNode();
		if (lotNode) {
			removeSerialDetails(lotNode);
			for (i = 1; i <= rows; i += 1) {
				detail = createElement("SerialDetails");
				setAttr(detail, "SerNo", fieldValue("hSerial" + i));
				setAttr(detail, "Qty", fieldValue("txtQty" + i));
				lotNode.appendChild(detail);
			}
		}

		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
